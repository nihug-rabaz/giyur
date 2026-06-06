const http = require("http");
const https = require("https");
const fs = require("fs");
const { URL } = require("url");

// Loads KEY=VALUE pairs from a .env file into process.env without overriding existing vars.
function loadEnv(path = `${__dirname}/.env`) {
  if (!fs.existsSync(path)) return;
  for (const line of fs.readFileSync(path, "utf8").split(/\r?\n/)) {
    if (!line.trim() || line.trim().startsWith("#")) continue;
    const match = line.match(/^\s*([\w.-]+)\s*=\s*(.*)\s*$/);
    if (!match) continue;
    let value = match[2];
    if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
      value = value.slice(1, -1);
    }
    if (!(match[1] in process.env)) process.env[match[1]] = value;
  }
}

class Config {
  // Reads and normalizes every setting the backend needs from environment variables.
  constructor(env) {
    this.siteUrl = (env.SP_SITE_URL || "").replace(/\/$/, "");
    this.listName = env.SP_LIST_NAME || "";
    this.itemType = env.SP_ITEM_TYPE || "";
    this.cookie = env.SP_COOKIE || "";
    this.username = env.SP_USERNAME || "";
    this.password = env.SP_PASSWORD || "";
    this.domain = env.SP_DOMAIN || "";
    this.workstation = env.SP_WORKSTATION || "";
    this.port = Number(env.PORT || 3000);
    this.allowOrigin = env.CORS_ORIGIN || "*";
    this.insecureTls = env.SP_INSECURE_TLS === "1";
  }

  // Picks NTLM when a service account is supplied, otherwise falls back to cookie reuse.
  get authMode() {
    return this.username && this.password ? "ntlm" : "cookie";
  }

  // Returns a config copy whose target site/list are overridden by the request (extension config).
  withTarget(siteUrl, listName) {
    const clone = Object.create(Config.prototype);
    Object.assign(clone, this, {
      siteUrl: siteUrl ? siteUrl.replace(/\/$/, "") : this.siteUrl,
      listName: listName || this.listName,
    });
    return clone;
  }
}

class AuthError extends Error {
  constructor(message) {
    super(message);
    this.status = 401;
  }
}

class SharePointError extends Error {
  constructor(status, body) {
    super(SharePointError.extract(body) || `SharePoint responded with ${status}`);
    this.status = status;
  }

  // Pulls the human-readable message out of a verbose SharePoint error payload.
  static extract(body) {
    try {
      const value = JSON.parse(body)?.error?.message?.value;
      return typeof value === "string" ? value : null;
    } catch {
      return null;
    }
  }
}

// Sends an HTTP(S) request with Node's built-in modules and resolves the raw response.
function rawRequest(method, url, headers, body, insecureTls) {
  return new Promise((resolve, reject) => {
    const target = new URL(url);
    const lib = target.protocol === "http:" ? http : https;
    const data = body == null ? null : typeof body === "string" ? body : JSON.stringify(body);
    const requestHeaders = { ...headers };
    if (data) requestHeaders["Content-Length"] = Buffer.byteLength(data);
    else if (method === "POST" || method === "PUT") requestHeaders["Content-Length"] = 0;

    const req = lib.request(
      {
        method,
        hostname: target.hostname,
        port: target.port || (target.protocol === "http:" ? 80 : 443),
        path: target.pathname + target.search,
        headers: requestHeaders,
        rejectUnauthorized: !insecureTls,
      },
      (res) => {
        let chunks = "";
        res.on("data", (chunk) => (chunks += chunk));
        res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body: chunks }));
      }
    );
    req.on("error", reject);
    if (data) req.write(data);
    req.end();
  });
}

class CookieAuth {
  // Reuses session cookies (per-request from the extension, or a static one from env).
  constructor(config, cookie) {
    this.config = config;
    this.cookie = cookie || config.cookie;
  }

  request(method, url, headers, body) {
    return rawRequest(method, url, { ...headers, Cookie: this.cookie }, body, this.config.insecureTls);
  }
}

class NtlmAuth {
  // Authenticates a service account through the NTLM handshake via the optional httpntlm library.
  constructor(config) {
    this.config = config;
    try {
      this.httpntlm = require("httpntlm");
    } catch {
      throw new Error("NTLM auth needs the 'httpntlm' package. Run: npm install httpntlm");
    }
  }

  request(method, url, headers, body) {
    const options = {
      url,
      headers,
      body: body == null ? "" : typeof body === "string" ? body : JSON.stringify(body),
      username: this.config.username,
      password: this.config.password,
      domain: this.config.domain,
      workstation: this.config.workstation,
      rejectUnauthorized: !this.config.insecureTls,
    };
    return new Promise((resolve, reject) => {
      this.httpntlm[method.toLowerCase()](options, (err, res) => {
        if (err) return reject(err);
        resolve({ status: res.statusCode, headers: res.headers, body: res.body });
      });
    });
  }
}

class SharePointClient {
  constructor(config, auth) {
    this.config = config;
    this.auth = auth;
    this.itemType = config.itemType || null;
    this.digest = null;
    this.digestExpiresAt = 0;
  }

  // Returns a cached form digest, fetching a fresh one from /_api/contextinfo when missing or stale.
  async getDigest(force = false) {
    if (!force && this.digest && Date.now() < this.digestExpiresAt) return this.digest;
    const res = await this.auth.request("POST", `${this.config.siteUrl}/_api/contextinfo`, {
      Accept: "application/json;odata=verbose",
      "Content-Type": "application/json;odata=verbose",
    });
    if (res.status === 401) throw new AuthError("Unauthorized while requesting contextinfo");
    if (res.status >= 400) throw new SharePointError(res.status, res.body);
    const info = SharePointClient.parse(res.body)?.d?.GetContextWebInformation;
    if (!info || !info.FormDigestValue) throw new SharePointError(res.status, "FormDigestValue missing");
    this.digest = info.FormDigestValue;
    this.digestExpiresAt = Date.now() + (info.FormDigestTimeoutSeconds || 1800) * 1000 - 60000;
    return this.digest;
  }

  // Resolves and caches the list entity type required for the __metadata.type field.
  async getItemType() {
    if (this.itemType) return this.itemType;
    const url = `${this._listUrl()}?$select=ListItemEntityTypeFullName`;
    const res = await this.auth.request("GET", url, { Accept: "application/json;odata=verbose" });
    if (res.status === 401) throw new AuthError("Unauthorized while reading list metadata");
    if (res.status >= 400) throw new SharePointError(res.status, res.body);
    const type = SharePointClient.parse(res.body)?.d?.ListItemEntityTypeFullName;
    if (!type) throw new SharePointError(res.status, "ListItemEntityTypeFullName not found");
    this.itemType = type;
    return type;
  }

  // Creates a list item, refreshing the digest and retrying once if it was rejected as stale.
  async createItem(fields) {
    const payload = JSON.stringify({ __metadata: { type: await this.getItemType() }, ...fields });
    let res = await this._postItem(payload, await this.getDigest());
    if (res.status === 403) res = await this._postItem(payload, await this.getDigest(true));
    if (res.status === 401) throw new AuthError("Unauthorized while creating item");
    if (res.status >= 400) throw new SharePointError(res.status, res.body);
    return SharePointClient.parse(res.body)?.d ?? {};
  }

  _postItem(payload, digest) {
    return this.auth.request(
      "POST",
      `${this._listUrl()}/items`,
      {
        Accept: "application/json;odata=verbose",
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest,
      },
      payload
    );
  }

  _listUrl() {
    return `${this.config.siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(this.config.listName)}')`;
  }

  static parse(body) {
    try {
      return JSON.parse(body);
    } catch {
      return null;
    }
  }
}

// Reads and JSON-parses a request body with a size guard.
function readBody(req, limit = 1e6) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk) => {
      data += chunk;
      if (data.length > limit) reject(new Error("Payload too large"));
    });
    req.on("end", () => {
      if (!data) return resolve({});
      try {
        resolve(JSON.parse(data));
      } catch {
        reject(new Error("Invalid JSON body"));
      }
    });
    req.on("error", reject);
  });
}

class ApiServer {
  constructor(config) {
    this.config = config;
  }

  start() {
    http.createServer((req, res) => this._handle(req, res)).listen(this.config.port, () => {
      console.log(`Backend listening on http://localhost:${this.config.port} (auth: ${this.config.authMode})`);
    });
  }

  // Routes requests; only POST /create-item and CORS preflight are served.
  async _handle(req, res) {
    this._cors(res);
    if (req.method === "OPTIONS") return this._send(res, 204, null);
    if (req.method === "POST" && req.url === "/create-item") return this._createItem(req, res);
    return this._send(res, 404, { success: false, error: "Not found" });
  }

  // Validates input, picks a client for the request's auth + target, and returns a clean JSON envelope.
  async _createItem(req, res) {
    try {
      const { title, fields, cookie, siteUrl, listName } = await readBody(req);
      const data = { ...(fields && typeof fields === "object" ? fields : {}) };
      if (title) data.Title = title;
      if (!Object.keys(data).length) return this._send(res, 400, { success: false, error: "Provide 'title' or 'fields'" });

      const config = this.config.withTarget(siteUrl, listName);
      if (!config.siteUrl || !config.listName) {
        return this._send(res, 400, { success: false, error: "siteUrl and listName are required (send them or set SP_SITE_URL/SP_LIST_NAME)" });
      }
      const client = this._clientFor(config, cookie);
      if (!client) {
        return this._send(res, 400, { success: false, error: "No cookie available: send 'cookie' in the request or set SP_COOKIE" });
      }
      const created = await client.createItem(data);
      this._send(res, 200, { success: true, data: { id: created.Id ?? created.ID ?? null } });
    } catch (err) {
      const status = err.status >= 400 && err.status < 600 ? err.status : 500;
      this._send(res, status, { success: false, error: err.message });
    }
  }

  // Builds a client for the chosen auth against the resolved target config.
  _clientFor(config, cookie) {
    if (config.authMode === "ntlm") return new SharePointClient(config, new NtlmAuth(config));
    const effective = cookie || config.cookie;
    if (!effective) return null;
    return new SharePointClient(config, new CookieAuth(config, effective));
  }

  _cors(res) {
    res.setHeader("Access-Control-Allow-Origin", this.config.allowOrigin);
    res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
    res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  }

  _send(res, status, payload) {
    if (payload === null) return res.writeHead(status).end();
    res.writeHead(status, { "Content-Type": "application/json" });
    res.end(JSON.stringify(payload));
  }
}

function main() {
  loadEnv();
  new ApiServer(new Config(process.env)).start();
}

main();
