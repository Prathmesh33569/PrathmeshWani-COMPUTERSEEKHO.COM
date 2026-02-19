// API Configuration
export const API_BASE_URL = "http://localhost:8080/api/enquiries";

// Logged-in staff ID (Replace with actual auth)
const staffData = JSON.parse(localStorage.getItem("staff"));

export const LOGGED_IN_STAFF_ID = staffData?.staffId;

// Enquiry Sources
export const ENQUIRY_SOURCES = [
  "Telephonic",
  "Walk-in",
  "Email",
  "Online",
  "Fax",
];

// Closure Reasons
export const CLOSURE_REASONS = [
  { id: 1, label: "Not Interested" },
  { id: 2, label: "Fee Issue" },
  { id: 3, label: "Joined Another Institute" },
  { id: 4, label: "Other" },
];

// Filter Options
export const FILTER_OPTIONS = {
  ALL: "all",
  TODAY: "today",
  PENDING: "pending",
};

// View Types
export const VIEW_TYPES = {
  MY_FOLLOWUPS: "my-followups",
  ALL_FOLLOWUPS: "all-followups",
};
