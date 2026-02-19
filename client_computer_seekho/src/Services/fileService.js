import axios from "axios";

// Create a separate axios instance for file uploads if needed, 
// or use the base URL from generic axios configuration if available.
// Assuming the backend is at localhost:8080 based on other files.
const API_URL = "http://localhost:8080/api";

const uploadImage = async (file) => {
  const formData = new FormData();
  formData.append("file", file);

  try {
    const response = await axios.post(`${API_URL}/upload`, formData, {
      headers: {
        "Content-Type": "multipart/form-data",
      },
    });
    return response.data; // Expected { imagePath: "/uploads/filename.ext" }
  } catch (error) {
    throw error;
  }
};

export default {
  uploadImage,
};
