// import React, { useState } from "react";
// import { useNavigate, Link } from "react-router-dom";
// import authService from "../../Services/authService.js";

// function Signup() {
//   const navigate = useNavigate();
//   const [formData, setFormData] = useState({
//     staffName: "",
//     staffMobile: "",
//     staffEmail: "",
//     staffUsername: "",
//     staffPassword: "",
//     staffRole: "teaching",
//     staffDesignation: "",
//     staffBio: "",
//   });

//   // IMAGE STATES 🖼️
//   const [staffImage, setStaffImage] = useState(null);
//   const [imagePreview, setImagePreview] = useState(null);

//   const [error, setError] = useState("");
//   const [success, setSuccess] = useState("");
//   const [loading, setLoading] = useState(false);

//   const handleChange = (e) => {
//     setFormData({
//       ...formData,
//       [e.target.name]: e.target.value,
//     });
//   };

//   // 📸 HANDLE IMAGE SELECTION
//   const handleImageChange = (e) => {
//     const file = e.target.files[0];

//     if (file) {
//       // Validate file type
//       const validTypes = ["image/jpeg", "image/jpg", "image/png"];
//       if (!validTypes.includes(file.type)) {
//         setError("Please select a valid image (JPG, JPEG, or PNG)");
//         return;
//       }

//       // Validate file size (5MB max)
//       const maxSize = 5 * 1024 * 1024; // 5MB in bytes
//       if (file.size > maxSize) {
//         setError("Image size should be less than 5MB");
//         return;
//       }

//       // Set image and create preview
//       setStaffImage(file);
//       setImagePreview(URL.createObjectURL(file));
//       setError(""); // Clear any previous errors
//     }
//   };

//   // 🗑️ REMOVE IMAGE
//   const handleRemoveImage = () => {
//     setStaffImage(null);
//     setImagePreview(null);
//     // Reset file input
//     document.getElementById("staffImage").value = "";
//   };

//   const handleSubmit = async (e) => {
//     e.preventDefault();
//     setError("");
//     setSuccess("");
//     setLoading(true);

//     try {
//       const response = await authService.signup(formData, staffImage);

//       console.log("Signup successful:", response);
//       setSuccess(response.message);

//       setTimeout(() => {
//         navigate("/admin/login");
//       }, 2000);
//     } catch (err) {
//       console.error("Signup error:", err);
//       setError(
//         err.response?.data?.message || "Registration failed. Please try again.",
//       );
//     } finally {
//       setLoading(false);
//     }
//   };

//   return (
//     <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-purple-600 via-purple-700 to-indigo-800 px-4 py-8">
//       <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-4xl animate-fadeIn">
//         <div className="text-center mb-8">
//           <h2 className="text-3xl font-bold text-gray-800 mb-2">
//             Staff Registration
//           </h2>
//           <p className="text-gray-600 text-sm">Join Computer Seekho Team</p>
//         </div>

//         {error && (
//           <div className="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 mb-6 rounded">
//             <p className="text-sm">{error}</p>
//           </div>
//         )}

//         {success && (
//           <div className="bg-green-50 border-l-4 border-green-500 text-green-700 p-4 mb-6 rounded">
//             <p className="text-sm">{success}</p>
//           </div>
//         )}

//         <form onSubmit={handleSubmit} className="space-y-6">
//           {/* 🖼️ IMAGE UPLOAD SECTION */}
//           <div className="flex flex-col items-center mb-6">
//             <label className="block text-sm font-semibold text-gray-700 mb-3">
//               Profile Picture
//             </label>

//             {/* Image Preview */}
//             {imagePreview ? (
//               <div className="relative">
//                 <img
//                   src={imagePreview}
//                   alt="Preview"
//                   className="w-32 h-32 rounded-full object-cover border-4 border-purple-500 shadow-lg"
//                 />
//                 <button
//                   type="button"
//                   onClick={handleRemoveImage}
//                   disabled={loading}
//                   className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-8 h-8 flex items-center justify-center hover:bg-red-600 transition shadow-lg disabled:opacity-50"
//                 >
//                   ✕
//                 </button>
//               </div>
//             ) : (
//               <div className="w-32 h-32 rounded-full bg-gray-200 flex items-center justify-center border-4 border-dashed border-gray-400">
//                 <svg
//                   className="w-12 h-12 text-gray-400"
//                   fill="none"
//                   stroke="currentColor"
//                   viewBox="0 0 24 24"
//                 >
//                   <path
//                     strokeLinecap="round"
//                     strokeLinejoin="round"
//                     strokeWidth="2"
//                     d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"
//                   />
//                 </svg>
//               </div>
//             )}

//             {/* File Input */}
//             <input
//               type="file"
//               id="staffImage"
//               name="staffImage"
//               accept="image/jpeg,image/jpg,image/png"
//               onChange={handleImageChange}
//               disabled={loading}
//               className="hidden"
//             />

//             <label
//               htmlFor="staffImage"
//               className="mt-4 px-6 py-2 bg-purple-100 text-purple-700 rounded-lg cursor-pointer hover:bg-purple-200 transition font-semibold text-sm disabled:opacity-50"
//             >
//               {imagePreview ? "Change Photo" : "Upload Photo"}
//             </label>

//             <p className="text-xs text-gray-500 mt-2">
//               JPG, JPEG, or PNG (Max 5MB)
//             </p>
//           </div>

//           {/* FORM FIELDS */}
//           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
//             <div>
//               <label
//                 htmlFor="staffName"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Full Name *
//               </label>
//               <input
//                 type="text"
//                 id="staffName"
//                 name="staffName"
//                 value={formData.staffName}
//                 onChange={handleChange}
//                 placeholder="Enter full name"
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               />
//             </div>

//             <div>
//               <label
//                 htmlFor="staffMobile"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Mobile Number *
//               </label>
//               <input
//                 type="tel"
//                 id="staffMobile"
//                 name="staffMobile"
//                 value={formData.staffMobile}
//                 onChange={handleChange}
//                 placeholder="Enter mobile number"
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               />
//             </div>
//           </div>

//           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
//             <div>
//               <label
//                 htmlFor="staffEmail"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Email *
//               </label>
//               <input
//                 type="email"
//                 id="staffEmail"
//                 name="staffEmail"
//                 value={formData.staffEmail}
//                 onChange={handleChange}
//                 placeholder="Enter email address"
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               />
//             </div>

//             <div>
//               <label
//                 htmlFor="staffUsername"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Username *
//               </label>
//               <input
//                 type="text"
//                 id="staffUsername"
//                 name="staffUsername"
//                 value={formData.staffUsername}
//                 onChange={handleChange}
//                 placeholder="Choose username"
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               />
//             </div>
//           </div>

//           <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
//             <div>
//               <label
//                 htmlFor="staffPassword"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Password *
//               </label>
//               <input
//                 type="password"
//                 id="staffPassword"
//                 name="staffPassword"
//                 value={formData.staffPassword}
//                 onChange={handleChange}
//                 placeholder="Create password (min 6 chars)"
//                 minLength="6"
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               />
//             </div>

//             <div>
//               <label
//                 htmlFor="staffRole"
//                 className="block text-sm font-semibold text-gray-700 mb-2"
//               >
//                 Staff Role *
//               </label>
//               <select
//                 id="staffRole"
//                 name="staffRole"
//                 value={formData.staffRole}
//                 onChange={handleChange}
//                 required
//                 disabled={loading}
//                 className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//               >
//                 <option value="teaching">Teaching</option>
//                 <option value="non-teaching">Non-Teaching</option>
//               </select>
//             </div>
//           </div>

//           <div>
//             <label
//               htmlFor="staffDesignation"
//               className="block text-sm font-semibold text-gray-700 mb-2"
//             >
//               Designation
//             </label>
//             <input
//               type="text"
//               id="staffDesignation"
//               name="staffDesignation"
//               value={formData.staffDesignation}
//               onChange={handleChange}
//               placeholder="e.g., Professor, Admin Officer"
//               disabled={loading}
//               className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
//             />
//           </div>

//           <div>
//             <label
//               htmlFor="staffBio"
//               className="block text-sm font-semibold text-gray-700 mb-2"
//             >
//               Bio
//             </label>
//             <textarea
//               id="staffBio"
//               name="staffBio"
//               value={formData.staffBio}
//               onChange={handleChange}
//               placeholder="Brief bio or description"
//               rows="3"
//               disabled={loading}
//               className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100 resize-none"
//             />
//           </div>

//           <button
//             type="submit"
//             disabled={loading}
//             className="w-full bg-gradient-to-r from-purple-600 to-indigo-600 text-white font-semibold py-3 rounded-lg hover:from-purple-700 hover:to-indigo-700 transition transform hover:scale-[1.02] disabled:opacity-60 disabled:cursor-not-allowed disabled:transform-none shadow-lg"
//           >
//             {loading ? "Registering..." : "Register"}
//           </button>
//         </form>

//         <p className="text-center text-gray-600 text-sm mt-6">
//           Already have an account?{" "}
//           <Link
//             to="/admin/login"
//             className="text-purple-600 font-semibold hover:text-purple-700 transition"
//           >
//             Login here
//           </Link>
//         </p>
//       </div>
//     </div>
//   );
// }

// export default Signup;

import React, { useState } from "react";
import { useNavigate, Link } from "react-router-dom";
import authService from "../../Services/authService.js";

function Signup() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    staffName: "",
    staffMobile: "",
    staffEmail: "",
    staffUsername: "",
    staffPassword: "",
    staffRole: "teaching",
    staffDesignation: "",
    staffBio: "",
  });

  // IMAGE STATES 🖼️
  const [staffImage, setStaffImage] = useState(null);
  const [imagePreview, setImagePreview] = useState(null);

  const [error, setError] = useState("");
  const [success, setSuccess] = useState("");
  const [loading, setLoading] = useState(false);

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  // 📸 HANDLE IMAGE SELECTION
  const handleImageChange = (e) => {
    const file = e.target.files[0];

    if (file) {
      // Validate file type
      const validTypes = ["image/jpeg", "image/jpg", "image/png"];
      if (!validTypes.includes(file.type)) {
        setError("Please select a valid image (JPG, JPEG, or PNG)");
        return;
      }

      // Validate file size (5MB max)
      const maxSize = 5 * 1024 * 1024; // 5MB in bytes
      if (file.size > maxSize) {
        setError("Image size should be less than 5MB");
        return;
      }

      // Set image and create preview
      setStaffImage(file);
      setImagePreview(URL.createObjectURL(file));
      setError(""); // Clear any previous errors
    }
  };

  // 🗑️ REMOVE IMAGE
  const handleRemoveImage = () => {
    setStaffImage(null);
    setImagePreview(null);
    // Reset file input
    const fileInput = document.getElementById("staffImage");
    if (fileInput) {
      fileInput.value = "";
    }
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setSuccess("");
    setLoading(true);

    // Validate required fields
    if (
      !formData.staffName ||
      !formData.staffMobile ||
      !formData.staffEmail ||
      !formData.staffUsername ||
      !formData.staffPassword
    ) {
      setError("Please fill in all required fields");
      setLoading(false);
      return;
    }

    // Validate password length
    if (formData.staffPassword.length < 6) {
      setError("Password must be at least 6 characters long");
      setLoading(false);
      return;
    }

    try {
      console.log("=== SIGNUP REQUEST ===");
      console.log("Form Data:", formData);
      console.log("Has Image:", !!staffImage);

      const response = await authService.signup(formData, staffImage);

      console.log("Signup successful:", response);
      setSuccess(
        response.message || "Registration successful! Redirecting to login...",
      );

      // Clear form
      setFormData({
        staffName: "",
        staffMobile: "",
        staffEmail: "",
        staffUsername: "",
        staffPassword: "",
        staffRole: "teaching",
        staffDesignation: "",
        staffBio: "",
      });
      setStaffImage(null);
      setImagePreview(null);

      setTimeout(() => {
        navigate("/admin/login");
      }, 2000);
    } catch (err) {
      console.error("Signup error:", err);
      console.error("Error response:", err.response);

      // Handle different error scenarios
      if (err.response?.data?.message) {
        setError(err.response.data.message);
      } else if (err.response?.data) {
        setError(err.response.data);
      } else if (err.message) {
        setError(err.message);
      } else {
        setError("Registration failed. Please try again.");
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-purple-600 via-purple-700 to-indigo-800 px-4 py-8">
      <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-4xl animate-fadeIn">
        <div className="text-center mb-8">
          <h2 className="text-3xl font-bold text-gray-800 mb-2">
            Staff Registration
          </h2>
          <p className="text-gray-600 text-sm">Join Computer Seekho Team</p>
        </div>

        {error && (
          <div className="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 mb-6 rounded">
            <p className="text-sm">{error}</p>
          </div>
        )}

        {success && (
          <div className="bg-green-50 border-l-4 border-green-500 text-green-700 p-4 mb-6 rounded">
            <p className="text-sm">{success}</p>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-6">
          {/* 🖼️ IMAGE UPLOAD SECTION */}
          <div className="flex flex-col items-center mb-6">
            <label className="block text-sm font-semibold text-gray-700 mb-3">
              Profile Picture
            </label>

            {/* Image Preview */}
            {imagePreview ? (
              <div className="relative">
                <img
                  src={imagePreview}
                  alt="Preview"
                  className="w-32 h-32 rounded-full object-cover border-4 border-purple-500 shadow-lg"
                />
                <button
                  type="button"
                  onClick={handleRemoveImage}
                  disabled={loading}
                  className="absolute -top-2 -right-2 bg-red-500 text-white rounded-full w-8 h-8 flex items-center justify-center hover:bg-red-600 transition shadow-lg disabled:opacity-50"
                >
                  ✕
                </button>
              </div>
            ) : (
              <div className="w-32 h-32 rounded-full bg-gray-200 flex items-center justify-center border-4 border-dashed border-gray-400">
                <svg
                  className="w-12 h-12 text-gray-400"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z"
                  />
                </svg>
              </div>
            )}

            {/* File Input */}
            <input
              type="file"
              id="staffImage"
              name="staffImage"
              accept="image/jpeg,image/jpg,image/png"
              onChange={handleImageChange}
              disabled={loading}
              className="hidden"
            />

            <label
              htmlFor="staffImage"
              className="mt-4 px-6 py-2 bg-purple-100 text-purple-700 rounded-lg cursor-pointer hover:bg-purple-200 transition font-semibold text-sm disabled:opacity-50"
            >
              {imagePreview ? "Change Photo" : "Upload Photo"}
            </label>

            <p className="text-xs text-gray-500 mt-2">
              JPG, JPEG, or PNG (Max 5MB)
            </p>
          </div>

          {/* FORM FIELDS */}
          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label
                htmlFor="staffName"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Full Name *
              </label>
              <input
                type="text"
                id="staffName"
                name="staffName"
                value={formData.staffName}
                onChange={handleChange}
                placeholder="Enter full name"
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              />
            </div>

            <div>
              <label
                htmlFor="staffMobile"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Mobile Number *
              </label>
              <input
                type="tel"
                id="staffMobile"
                name="staffMobile"
                value={formData.staffMobile}
                onChange={handleChange}
                placeholder="Enter mobile number"
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label
                htmlFor="staffEmail"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Email *
              </label>
              <input
                type="email"
                id="staffEmail"
                name="staffEmail"
                value={formData.staffEmail}
                onChange={handleChange}
                placeholder="Enter email address"
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              />
            </div>

            <div>
              <label
                htmlFor="staffUsername"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Username *
              </label>
              <input
                type="text"
                id="staffUsername"
                name="staffUsername"
                value={formData.staffUsername}
                onChange={handleChange}
                placeholder="Choose username"
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              />
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
            <div>
              <label
                htmlFor="staffPassword"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Password *
              </label>
              <input
                type="password"
                id="staffPassword"
                name="staffPassword"
                value={formData.staffPassword}
                onChange={handleChange}
                placeholder="Create password (min 6 chars)"
                minLength="6"
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              />
            </div>

            <div>
              <label
                htmlFor="staffRole"
                className="block text-sm font-semibold text-gray-700 mb-2"
              >
                Staff Role *
              </label>
              <select
                id="staffRole"
                name="staffRole"
                value={formData.staffRole}
                onChange={handleChange}
                required
                disabled={loading}
                className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
              >
                <option value="teaching">Teaching</option>
                <option value="non-teaching">Non-Teaching</option>
              </select>
            </div>
          </div>

          <div>
            <label
              htmlFor="staffDesignation"
              className="block text-sm font-semibold text-gray-700 mb-2"
            >
              Designation
            </label>
            <input
              type="text"
              id="staffDesignation"
              name="staffDesignation"
              value={formData.staffDesignation}
              onChange={handleChange}
              placeholder="e.g., Professor, Admin Officer"
              disabled={loading}
              className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100"
            />
          </div>

          <div>
            <label
              htmlFor="staffBio"
              className="block text-sm font-semibold text-gray-700 mb-2"
            >
              Bio
            </label>
            <textarea
              id="staffBio"
              name="staffBio"
              value={formData.staffBio}
              onChange={handleChange}
              placeholder="Brief bio or description"
              rows="3"
              disabled={loading}
              className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100 resize-none"
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-gradient-to-r from-purple-600 to-indigo-600 text-white font-semibold py-3 rounded-lg hover:from-purple-700 hover:to-indigo-700 transition transform hover:scale-[1.02] disabled:opacity-60 disabled:cursor-not-allowed disabled:transform-none shadow-lg"
          >
            {loading ? "Registering..." : "Register"}
          </button>
        </form>

        <p className="text-center text-gray-600 text-sm mt-6">
          Already have an account?{" "}
          <Link
            to="/admin/login"
            className="text-purple-600 font-semibold hover:text-purple-700 transition"
          >
            Login here
          </Link>
        </p>
      </div>
    </div>
  );
}

export default Signup;
