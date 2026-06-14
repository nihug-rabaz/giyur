const SHAREPOINT_LOOKUP = {
  siteUrl: "https://rabaz.army.idf/sites/Giur/Giuron",
  listTitle: "ניהול מתגיירים",
  join: {
    tableColumnDisplay: "שיוך לתהליך (מ.א מתגייר/ת)",
    listFieldDisplay: "שם התהליך (מ.א מתגייר/ת)",
    listFieldInternal: "Title",
  },
};

const SHAREPOINT_FIELD_MAP = {
  "שם משפחה": "_x05e9__x05dd__x0020__x05de__x05",
  "שם פרטי": "_x05e9__x05dd__x0020__x05e4__x05",
};

window.SHAREPOINT_LOOKUP = SHAREPOINT_LOOKUP;
window.SHAREPOINT_FIELD_MAP = SHAREPOINT_FIELD_MAP;
