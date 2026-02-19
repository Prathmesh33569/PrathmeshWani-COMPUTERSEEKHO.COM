import api from "./api";

/**
 * Get all teaching staff (faculty) with their images
 */
export const getFacultyList = async () => {
  try {
    const response = await api.get("/staff/faculty");
    return response.data;
  } catch (error) {
    console.error("Error fetching faculty:", error);
    throw error;
  }
};

/**
 * Get all staff (teaching + non-teaching)
 */
export const getAllStaff = async () => {
  try {
    const response = await api.get("/staff/all");
    return response.data;
  } catch (error) {
    console.error("Error fetching staff:", error);
    throw error;
  }
};

/**
 * Get staff by ID
 */
export const getStaffById = async (staffId) => {
  try {
    const response = await api.get(`/staff/${staffId}`);
    return response.data;
  } catch (error) {
    console.error("Error fetching staff:", error);
    throw error;
  }
};