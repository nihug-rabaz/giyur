class PlaceholderMapper {
  static ALIASES = [
    ["תחילה", "התחלה", "תחלה", "start"],
    ["סיום", "סיים", "end"],
    ["אירוע", "event", "שם"],
  ];

  static autoMatch(placeholder, headers) {
    const target = this._normalize(placeholder);

    const exact = headers.find((h) => this._normalize(h.label) === target);
    if (exact) return exact.id;

    const contains = headers.find((h) => {
      const label = this._normalize(h.label);
      return label.includes(target) || target.includes(label);
    });
    if (contains) return contains.id;

    const aliasGroup = PlaceholderMapper.ALIASES.find((group) =>
      group.some((word) => target.includes(this._normalize(word)))
    );
    if (aliasGroup) {
      const aliasMatch = headers.find((h) => {
        const label = this._normalize(h.label);
        return aliasGroup.some((word) => label.includes(this._normalize(word)));
      });
      if (aliasMatch) return aliasMatch.id;
    }

    return "";
  }

  static buildAutoMapping(placeholders, headers) {
    const mapping = {};
    placeholders.forEach((p) => {
      mapping[p] = this.autoMatch(p, headers);
    });
    return mapping;
  }

  static _normalize(text) {
    return (text || "")
      .toLowerCase()
      .replace(/[\s"'`.,:;()\[\]-]/g, "")
      .trim();
  }
}

window.PlaceholderMapper = PlaceholderMapper;
