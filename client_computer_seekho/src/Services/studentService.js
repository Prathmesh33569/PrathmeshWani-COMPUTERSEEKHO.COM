import axios from 'axios';

const API_BASE_URL = 'http://localhost:8080/api';

// Create axios instances
const studentApi = axios.create({
  baseURL: `${API_BASE_URL}/students`,
  headers: { 'Content-Type': 'application/json' }
});

const batchApi = axios.create({
  baseURL: `${API_BASE_URL}/batches`,
  headers: { 'Content-Type': 'application/json' }
});

const paymentApi = axios.create({
  baseURL: `${API_BASE_URL}/payments`,
  headers: { 'Content-Type': 'application/json' }
});

/**
 * Create payment (before student registration)
 */
export const createPayment = async (paymentData) => {
  try {
    const response = await paymentApi.post('/create', paymentData);
    return response.data;
  } catch (error) {
    console.error('Error creating payment:', error);
    throw error;
  }
};

/**
 * Register student from enquiry (after payment)
 */
export const registerStudent = async (registrationData) => {
  try {
    const response = await studentApi.post('/register', registrationData);
    return response.data;
  } catch (error) {
    console.error('Error registering student:', error);
    throw error;
  }
};

/**
 * Get all batches
 */
export const getAllBatches = async () => {
  try {
    const response = await batchApi.get('');
    return response.data;
  } catch (error) {
    console.error('Error fetching batches:', error);
    throw error;
  }
};

/**
 * Get active batches only
 */
export const getActiveBatches = async () => {
  try {
    const response = await batchApi.get('/active');
    return response.data;
  } catch (error) {
    console.error('Error fetching active batches:', error);
    throw error;
  }
};

/**
 * Get batches by course
 */
export const getBatchesByCourse = async (courseId) => {
  try {
    const response = await batchApi.get(`/course/${courseId}`);
    console.log('Batches from backend for courseId:', courseId, response.data);
    return response.data;
  } catch (error) {
    console.error('Error fetching batches by course:', error);
    throw error;
  }
};

/**
 * Get all registered students with complete details
 */
 export const getAllRegisteredStudents = async () => {
  try {
    const response = await studentApi.get('/registered');
    console.log('Registered students from backend:', response.data);
    return response.data;
  } catch (error) {
    console.error('Error fetching registered students:', error);
    throw error;
  }
};

/**
 * Get all payments
 */
export const getAllPayments = async () => {
  try {
    const response = await paymentApi.get('');
    return response.data;
  } catch (error) {
    console.error('Error fetching payments:', error);
    throw error;
  }
};

/**
 * Get payments by student
 */
export const getPaymentsByStudent = async (studentId) => {
  try {
    const response = await paymentApi.get(`/student/${studentId}`);
    return response.data;
  } catch (error) {
    console.error('Error fetching payments for student:', error);
    throw error;
  }
};

export default {
  createPayment,
  registerStudent,
  getAllBatches,
  getBatchesByCourse,
  getAllRegisteredStudents ,
  getAllPayments,
  getPaymentsByStudent
};