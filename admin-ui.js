// Presentation-only layer for the admin screen: tab switching + a vanilla "radioactive lab"
// molecule background. Touches no settings logic (admin.js owns all behaviour).

class LabTabs {
  constructor(navId) {
    this.nav = document.getElementById(navId);
    if (!this.nav) return;
    this.tabs = [...this.nav.querySelectorAll(".lab-tab")];
    this.panels = [...document.querySelectorAll("[data-panel]")];
    this.tabs.forEach((tab) => tab.addEventListener("click", () => this._activate(tab.dataset.target)));
  }

  _activate(targetId) {
    this.tabs.forEach((t) => t.classList.toggle("is-active", t.dataset.target === targetId));
    this.panels.forEach((p) => p.classList.toggle("is-active", p.id === targetId));
  }
}

// Floating atoms connected by bonds — drawn on a full-screen canvas behind the content.
class MoleculeField {
  constructor(canvasId) {
    this.canvas = document.getElementById(canvasId);
    if (!this.canvas) return;
    this.ctx = this.canvas.getContext("2d");
    this.atoms = [];
    this.linkDist = 150;
    this.reduced = window.matchMedia?.("(prefers-reduced-motion: reduce)").matches;

    this._resize();
    window.addEventListener("resize", () => this._resize());
    this._seed();
    this._loop();
  }

  _resize() {
    this.dpr = Math.min(window.devicePixelRatio || 1, 2);
    this.w = window.innerWidth;
    this.h = window.innerHeight;
    this.canvas.width = this.w * this.dpr;
    this.canvas.height = this.h * this.dpr;
    this.canvas.style.width = `${this.w}px`;
    this.canvas.style.height = `${this.h}px`;
    this.ctx.setTransform(this.dpr, 0, 0, this.dpr, 0, 0);
  }

  _seed() {
    const count = Math.max(26, Math.min(60, Math.round((this.w * this.h) / 30000)));
    const speed = this.reduced ? 0.08 : 0.35;
    for (let i = 0; i < count; i++) {
      this.atoms.push({
        x: Math.random() * this.w,
        y: Math.random() * this.h,
        vx: (Math.random() - 0.5) * speed,
        vy: (Math.random() - 0.5) * speed,
        r: 1.6 + Math.random() * 2.6,
        hot: Math.random() < 0.22,
      });
    }
  }

  _step() {
    this.atoms.forEach((a) => {
      a.x += a.vx;
      a.y += a.vy;
      if (a.x < 0 || a.x > this.w) a.vx *= -1;
      if (a.y < 0 || a.y > this.h) a.vy *= -1;
    });
  }

  _draw() {
    const ctx = this.ctx;
    ctx.clearRect(0, 0, this.w, this.h);

    for (let i = 0; i < this.atoms.length; i++) {
      for (let j = i + 1; j < this.atoms.length; j++) {
        const a = this.atoms[i];
        const b = this.atoms[j];
        const dx = a.x - b.x;
        const dy = a.y - b.y;
        const dist = Math.hypot(dx, dy);
        if (dist > this.linkDist) continue;
        const alpha = (1 - dist / this.linkDist) * 0.5;
        ctx.strokeStyle = `rgba(57, 255, 153, ${alpha})`;
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(a.x, a.y);
        ctx.lineTo(b.x, b.y);
        ctx.stroke();
      }
    }

    this.atoms.forEach((a) => {
      const color = a.hot ? "0, 229, 255" : "57, 255, 153";
      const glow = ctx.createRadialGradient(a.x, a.y, 0, a.x, a.y, a.r * 5);
      glow.addColorStop(0, `rgba(${color}, 0.9)`);
      glow.addColorStop(1, `rgba(${color}, 0)`);
      ctx.fillStyle = glow;
      ctx.beginPath();
      ctx.arc(a.x, a.y, a.r * 5, 0, Math.PI * 2);
      ctx.fill();

      ctx.fillStyle = `rgba(${color}, 1)`;
      ctx.beginPath();
      ctx.arc(a.x, a.y, a.r, 0, Math.PI * 2);
      ctx.fill();
    });
  }

  _loop() {
    this._step();
    this._draw();
    requestAnimationFrame(() => this._loop());
  }
}

new LabTabs("labTabs");
new MoleculeField("moleculeCanvas");
