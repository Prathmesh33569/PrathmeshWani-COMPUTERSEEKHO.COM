import axios from 'axios';

const API_BASE_URL = 'http://localhost:8080/api';

const receiptApi = axios.create({
  baseURL: `${API_BASE_URL}/receipts`,
  headers: { 'Content-Type': 'application/json' }
});

/**
 * Download receipt as PDF
 * TODO: Your friend will implement backend
 */
export const downloadReceipt = async (receiptId) => {
  try {
    const response = await receiptApi.get(`/${receiptId}/download`, {
      responseType: 'blob' // For PDF download
    });
    
    // Create download link
    const url = window.URL.createObjectURL(new Blob([response.data]));
    const link = document.createElement('a');
    link.href = url;
    link.setAttribute('download', `receipt_${receiptId}.pdf`);
    document.body.appendChild(link);
    link.click();
    link.remove();
    
    return response.data;
  } catch (error) {
    console.error('Error downloading receipt:', error);
    throw error;
  }
};

/**
 * Send receipt via email
 * TODO: Your friend will implement backend
 */
// export const sendReceipt = async (receiptId, email) => {
//   try {
//     const response = await receiptApi.post(`/${receiptId}/send`, null, {
//       params: { email }
//     });
//     return response.data;
//   } catch (error) {
//     console.error('Error sending receipt:', error);
//     throw error;
//   }
// };


export const sendReceipt = async (receiptId) => {
  try {
    const response = await axios.get(
      `${API_BASE_URL}/receipt/pdf/${receiptId}/email`
    );
    return response.data;
  } catch (error) {
    console.error("Error sending receipt:", error);
    throw error;
  }
};


/**
 * Get receipt details by ID
 * TODO: Your friend will implement backend
 */
export const getReceiptById = async (receiptId) => {
  try {
    const response = await receiptApi.get(`/${receiptId}`);
    return response.data;
  } catch (error) {
    console.error('Error fetching receipt:', error);
    throw error;
  }
};

export default {
  downloadReceipt,
  sendReceipt,
  getReceiptById
};