const SHAREPOINT_LOOKUP = {
  siteUrl: "https://www.molecule.co.il/legacy-deployed/max-app",
  listTitle: "רשימת בסיס",
  join: {
    tableColumnDisplay: "שיוך לתהליך",
    listFieldDisplay: "שם התהליך",
    listFieldInternal: "",
  },
};

const SHAREPOINT_FIELD_MAP = {
  "שם פרטי": "OData_x05e9x05ddx0020x05e4_x05",
  "שם משפחה": "OData_x05e9x05ddx0020x05de_x05",
};

window.SHAREPOINT_LOOKUP = SHAREPOINT_LOOKUP;
window.SHAREPOINT_FIELD_MAP = SHAREPOINT_FIELD_MAP;
