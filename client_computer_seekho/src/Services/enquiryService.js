import axios from 'axios';
import { API_BASE_URL } from '../utils/constants';

// Create axios instance
const api = axios.create({
  baseURL: API_BASE_URL,
  headers: {
    'Content-Type': 'application/json'
  }
});

// Course API (separate base URL)
const courseApi = axios.create({
  baseURL: 'http://localhost:8080/api/courses',
  headers: {
    'Content-Type': 'application/json'
  }
});

// API Service Functions

/**
 * Get all courses for dropdown
 * Uses existing /api/courses endpoint that returns CourseResponseDTO list
 */
export const getAllCourses = async () => {
  try {
    const response = await courseApi.get('');
    return response.data;
  } catch (error) {
    console.error('Error fetching courses:', error);
    throw error;
  }
};

/**
 * Get only active courses for enquiry dropdown (better option)
 */
export const getActiveCourses = async () => {
  try {
    const response = await courseApi.get('/active');
    return response.data;
  } catch (error) {
    console.error('Error fetching active courses:', error);
    throw error;
  }
};

/**
 * Get follow-ups for specific staff member
 */
export const getFollowupsForStaff = async (staffId) => {
  try {
    const response = await api.get(`/followups/staff/${staffId}`);
    return response.data;
  } catch (error) {
    console.error('Error fetching staff followups:', error);
    throw error;
  }
};

/**
 * Get all follow-ups (Admin view)
 */
export const getAllFollowups = async () => {
  try {
    const response = await api.get('/followups/all');
    return response.data;
  } catch (error) {
    console.error('Error fetching all followups:', error);
    throw error;
  }
};

/**
 * Create new enquiry
 */
export const createEnquiry = async (enquiryData) => {
  try {
    const response = await api.post('', enquiryData);
    return response.data;
  } catch (error) {
    console.error('Error creating enquiry:', error);
    throw error;
  }
};

/**
 * Update follow-up (CALL button action)
 */
export const updateFollowup = async (followupData) => {
  try {
    const response = await api.put('/followup', followupData);
    return response.data;
  } catch (error) {
    console.error('Error updating followup:', error);
    throw error;
  }
};

/**
 * Get single enquiry by ID
 */
export const getEnquiryById = async (enquiryId) => {
  try {
    const response = await api.get(`/${enquiryId}`);
    return response.data;
  } catch (error) {
    console.error('Error fetching enquiry:', error);
    throw error;
  }
};

/**
 * Update enquiry details
 */
export const updateEnquiry = async (enquiryId, enquiryData) => {
  try {
    const response = await api.put(`/${enquiryId}`, enquiryData);
    return response.data;
  } catch (error) {
    console.error('Error updating enquiry:', error);
    throw error;
  }
};

export default {
  getFollowupsForStaff,
  getAllFollowups,
  createEnquiry,
  updateFollowup,
  getEnquiryById,
  updateEnquiry,
  getAllCourses,
  getActiveCourses
};