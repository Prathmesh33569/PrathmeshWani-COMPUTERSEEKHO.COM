// import api from "./api";

// const authService = {
//   // Login
//   login: async (staffUsername, staffPassword) => {
//     const response = await api.post("/auth/login", {
//       staffUsername,
//       staffPassword,
//     });

//     if (response.data.token) {
//       // Store token and staff data in localStorage
//       localStorage.setItem("token", response.data.token);
//       localStorage.setItem("staff", JSON.stringify(response.data));
//     }

//     return response.data;
//   },

//   // 🚀 Signup WITH IMAGE UPLOAD (MULTIPART)
//   signup: async (signupData, staffImage) => {
//     // Create FormData for multipart/form-data
//     const formData = new FormData();

//     // Append all text fields
//     formData.append("staffName", signupData.staffName);
//     formData.append("staffMobile", signupData.staffMobile);
//     formData.append("staffEmail", signupData.staffEmail);
//     formData.append("staffUsername", signupData.staffUsername);
//     formData.append("staffPassword", signupData.staffPassword);
//     formData.append("staffRole", signupData.staffRole);
//     formData.append("staffDesignation", signupData.staffDesignation || "");
//     formData.append("staffBio", signupData.staffBio || "");

//     // Append image if exists
//     if (staffImage) {
//       formData.append("staffImage", staffImage);
//     }

//     // Send multipart request
//     const response = await api.post("/auth/signup", formData, {
//       headers: {
//         "Content-Type": "multipart/form-data",
//       },
//     });

//     return response.data;
//   },

//   // Logout
//   logout: () => {
//     localStorage.removeItem("token");
//     localStorage.removeItem("staff");
//   },

//   // Get current staff
//   getCurrentStaff: () => {
//     const staffData = localStorage.getItem("staff");
//     return staffData ? JSON.parse(staffData) : null;
//   },

//   // Check if logged in
//   isLoggedIn: () => {
//     return localStorage.getItem("token") !== null;
//   },
// };

// export default authService;

import api from "./api";

const authService = {
  // Login
  login: async (staffUsername, staffPassword) => {
    try {
      const response = await api.post("/auth/login", {
        staffUsername,
        staffPassword,
      });

      if (response.data.token) {
        // Store token and staff data in localStorage
        localStorage.setItem("token", response.data.token);
        localStorage.setItem("staff", JSON.stringify(response.data));
      }

      return response.data;
    } catch (error) {
      console.error("Login error:", error);
      throw error;
    }
  },

  // 🚀 Signup WITH IMAGE UPLOAD (MULTIPART)
  signup: async (signupData, staffImage) => {
    try {
      console.log("=== AUTH SERVICE: SIGNUP ===");
      console.log("Signup Data:", signupData);
      console.log("Staff Image:", staffImage);

      // Create FormData for multipart/form-data
      const formData = new FormData();

      // Append all text fields
      formData.append("staffName", signupData.staffName || "");
      formData.append("staffMobile", signupData.staffMobile || "");
      formData.append("staffEmail", signupData.staffEmail || "");
      formData.append("staffUsername", signupData.staffUsername || "");
      formData.append("staffPassword", signupData.staffPassword || "");
      formData.append("staffRole", signupData.staffRole || "teaching");
      formData.append("staffDesignation", signupData.staffDesignation || "");
      formData.append("staffBio", signupData.staffBio || "");

      // Append image if exists
      if (staffImage) {
        formData.append("staffImage", staffImage);
        console.log("Image appended to FormData:", staffImage.name);
      } else {
        console.log("No image to upload");
      }

      // Log FormData contents (for debugging)
      console.log("FormData entries:");
      for (let pair of formData.entries()) {
        console.log(pair[0], pair[1]);
      }

      // Send multipart request
      const response = await api.post("/auth/signup", formData, {
        headers: {
          "Content-Type": "multipart/form-data",
        },
      });

      console.log("Signup response:", response.data);
      return response.data;
    } catch (error) {
      console.error("Signup error in authService:", error);
      console.error("Error response:", error.response);
      throw error;
    }
  },

  // Logout
  logout: () => {
    localStorage.removeItem("token");
    localStorage.removeItem("staff");
  },

  // Get current staff
  getCurrentStaff: () => {
    const staffData = localStorage.getItem("staff");
    return staffData ? JSON.parse(staffData) : null;
  },

  // Check if logged in
  isLoggedIn: () => {
    return localStorage.getItem("token") !== null;
  },

  // Get token
  getToken: () => {
    return localStorage.getItem("token");
  },
};

export default authService;
