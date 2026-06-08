// ============================================================
// WIT Canada — Global Config
// Single source of truth for constants: sheet names, headers,
// status enums, identifiers (Drive folders, groups), and URLs.
// ============================================================

// Deployed web-app base URL. Both pages are served from this one
// endpoint, routed by the ?page= parameter (see WebApp.doGet).
const WEB_APP_URL =
  "https://script.google.com/a/macros/women-in-tech.org/s/AKfycbxh3-EQ7VUbbf1WMn9q9aSBBCnSimhRST5QwEjs6VDXij07JwJQMP0Md99DqpqrFNmU/exec";

const CONFIG = {
  SHEET: {
    MAIN: "WIT_Members",
    EXTERNAL: "WIT_External",
    HEADER_ROW: 2,
    GROUP_EMAIL_ROW: 3, // row holding group email addresses (GroupsView)
    DATA_START_ROW: 4,

    // Generated report sheets (rebuilt by the report builders;
    // also opened from the menu, hence centralized here).
    REPORTS: {
      EMAIL_REQUESTS: "Report_Email_Requests",
      PHOTO_BIO: "Report_Photos&Bios",
      GROUPS: "Report_Groups",
    },
  },

  STATUS: {
    NEW: "—",
    REQUEST: "requested",
    SENT: "sent",
    CREATED: "created",
    NOT_ACTIVATED: "not activated",
    ACTIVE: "active",
  },

  CONTRACT: {
    NOT_SENT: "not sent",
    SENT: "sent",
    SIGNED: "signed",
  },

  GROUPS_STATUS: {
    DEFAULT: "—",
    REQUESTED: "requested",
    ADDED: "added",
  },

  EMAIL: {
    OPS_LEAD: "yevheniia.shavrukova@women-in-tech.org",
    OPS_LEAD_NAME: "Jane",
  },

  MEMBER_STATUS: {
    DEFAULT: "—",
    ONBOARDING: "onboarding",
    ACTIVE: "active",
  },

  HEADERS: {
    FIRST_NAME: "First name",
    LAST_NAME: "Last name",
    ROLE: "Role",
    WIT_EMAIL: "WIT Email",
    PERSONAL_EMAIL: "Personal Email",
    EMAIL_STATUS: "Email Status",
    ADDED_TO_GROUPS: "Add to groups",
    INTRO_SENT: "Intro sent",
    PHOTO: "Photo",
    BIO: "Bio",
    CONTRACT_STATUS: "Contract status",
    MEMBER_STATUS: "Member Status",
    PHONE: "Phone Number (Whats App)",
    LINKEDIN: "LinkedIn URL",
    BIRTHDAY: "Birthday",
  },

  GROUPS: {
    HEADER_TO_EMAIL: {
      "Canada Team": "canada-team@women-in-tech.org",
      Leadership: "canada-leadership@women-in-tech.org",
      "NA Awards": "na-awards@women-in-tech.org",
      Marketing: "canada-marketing@women-in-tech.org",
      Partnerships: "canada-partnerships@women-in-tech.org",
      Mentorship: "canada-mentorship@women-in-tech.org",
      "Region-heads": "canada-region-heads@women-in-tech.org",
      Alberta: "canada-alberta@women-in-tech.org",
      BC: "canada-bc@women-in-tech.org",
      Maritimes: "canada-maritimes@women-in-tech.org",
      Ontario: "canada-ontario@women-in-tech.org",
      Quebec: "canada-quebec@women-in-tech.org",
    },
  },

  REGIONS: {
    ALBERTA: "Alberta",
    BC: "BC",
    MARITIMES: "Maritimes",
    ONTARIO: "Ontario",
    QUEBEC: "Quebec",
  },

  UI: {
    CONFIRMATION: true,
    ALERTS: true,
  },

  // Keys for values stashed in PropertiesService.
  PROPERTIES: {
    SYNC_REPORT: "lastSyncReport",
  },

  DRIVE: {
    PHOTO_FOLDER_ID: "1SdzoC6kMquUP14atedTNOGZouYQ-VG6t",
    BIO_FOLDER_ID: "1J90gz9dTXAuo6xfgiO9yeyh69zftVu9f",

    // Used by the Drive-structure export tool (tools/DriveStructure.js).
    STRUCTURE_OUTPUT_FOLDER_ID: "10Yzu4SKwqWiepB5-LfvGN0eRaHSgektQ",
    SHARED_DRIVES: [
      { id: "0ABdfrNcsjynoUk9PVA", name: "General" },
      { id: "0ADJyJDduRllQUk9PVA", name: "Leadership" },
    ],
  },

  URLS: {
    MEMBER_GUIDE: WEB_APP_URL,
    SIGNATURE_GENERATOR: `${WEB_APP_URL}?page=signature`,
  },
};

// ============================================================
// Derived config — defined after the literal so it can reference
// CONFIG. Safe because it lives in the same file (the load-order
// caveat only affects CONFIG references from OTHER files' top level).
// ============================================================

// Region display/iteration order: the five regions plus a catch-all.
CONFIG.REGION_ORDER = [...Object.values(CONFIG.REGIONS), "Unknown"];

// Status colors, keyed by Email-Status value. `bg` is used for sheet
// cell backgrounds; `bg`+`fg` together style the HTML email badges.
CONFIG.COLORS = {
  STATUS: {
    [CONFIG.STATUS.REQUEST]:       { bg: "#fff3cd", fg: "#856404" },
    [CONFIG.STATUS.SENT]:          { bg: "#cfe2ff", fg: "#0a4a90" },
    [CONFIG.STATUS.CREATED]:       { bg: "#d1e7dd", fg: "#0a5933" },
    [CONFIG.STATUS.NOT_ACTIVATED]: { bg: "#fde8e8", fg: "#9c1c1c" },
  },
};
