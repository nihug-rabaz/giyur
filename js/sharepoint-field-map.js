const SHAREPOINT_LOOKUP = {
  listUrl:
    "https://rabaz.army.idf/sites/Giur/Giuron/_api/web/Lists/getByTitle('%D7%A0%D7%99%D7%94%D7%95%D7%9C%20%D7%9E%D7%AA%D7%92%D7%99%D7%99%D7%A8%D7%99%D7%9D')/items",
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
