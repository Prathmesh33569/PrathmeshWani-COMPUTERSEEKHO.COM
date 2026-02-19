// import React from "react";

// const FacultyCard = ({ faculty }) => {
//   const { name, role, description, imageUrl } = faculty;

//   return (
//     <div className="bg-white rounded-2xl shadow-md overflow-hidden">
      
//       {/* IMAGE */}
//       <img
//         src={imageUrl}
//         alt={name}
//         className="w-full h-137.5 object-cover"
//       />

//       {/* CONTENT */}
//       <div className="p-6">
//         <h3 className="text-2xl font-bold text-gray-800">
//           {name}
//         </h3>

//         <p className="text-lg text-blue-600 font-medium mt-1"> 
//           {role}
//         </p>

//         <p className="text-gray-600 mt-4 leading-relaxed">
//           {description}
//         </p>
//       </div>
//     </div>
//   );
// };

// export default FacultyCard;

import React from "react";

function FacultyCard({ faculty }) {
  // BETTER image URL construction
  const getImageUrl = (photoUrl) => {
    if (!photoUrl) {
      // Default placeholder if no image
      return "https://ui-avatars.com/api/?name=" + encodeURIComponent(faculty.staffName) + "&size=200&background=7c3aed&color=fff";
    }
    
    // Remove leading slash if present and construct full URL
    const cleanPath = photoUrl.startsWith('/') ? photoUrl.substring(1) : photoUrl;
    const fullUrl = `http://localhost:8080/${cleanPath}`;
    
    console.log("🖼️ Image URL:", fullUrl); // Debug log
    return fullUrl;
  };

  return (
    <div className="bg-white rounded-xl shadow-lg overflow-hidden hover:shadow-2xl transition-shadow duration-300">
      <div className="md:flex">
        {/* Image Section */}
        <div className="md:w-1/3 bg-gradient-to-br from-purple-100 to-indigo-100 flex items-center justify-center p-8">
          <div className="relative">
            <img
              src={getImageUrl(faculty.photoUrl)}
              alt={faculty.staffName}
              className="w-48 h-48 rounded-full object-cover border-4 border-white shadow-xl"
              onError={(e) => {
                console.error("❌ Image failed to load:", faculty.photoUrl);
                // Fallback to avatar with initials
                e.target.src = "https://ui-avatars.com/api/?name=" + encodeURIComponent(faculty.staffName) + "&size=200&background=7c3aed&color=fff";
              }}
              onLoad={() => {
                console.log("✅ Image loaded successfully:", faculty.photoUrl);
              }}
            />
            {/* Loading indicator */}
            {!faculty.photoUrl && (
              <div className="absolute inset-0 flex items-center justify-center">
                <span className="text-xs text-gray-500">No Photo</span>
              </div>
            )}
          </div>
        </div>

        {/* Content Section */}
        <div className="md:w-2/3 p-8">
          <div className="flex items-center mb-4">
            <h3 className="text-2xl font-bold text-gray-800">
              {faculty.staffName}
            </h3>
            {faculty.staffRole === "teaching" && (
              <span className="ml-4 px-3 py-1 bg-purple-100 text-purple-700 rounded-full text-xs font-semibold">
                Faculty
              </span>
            )}
          </div>

          {faculty.staffDesignation && (
            <p className="text-purple-600 font-semibold mb-3">
              {faculty.staffDesignation}
            </p>
          )}

          {faculty.staffBio && (
            <p className="text-gray-600 leading-relaxed mb-4">
              {faculty.staffBio}
            </p>
          )}

          <div className="flex flex-wrap gap-4 text-sm text-gray-500">
            {faculty.staffEmail && (
              <div className="flex items-center">
                <svg
                  className="w-5 h-5 mr-2 text-purple-500"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"
                  />
                </svg>
                {faculty.staffEmail}
              </div>
            )}

            {faculty.staffMobile && (
              <div className="flex items-center">
                <svg
                  className="w-5 h-5 mr-2 text-purple-500"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M3 5a2 2 0 012-2h3.28a1 1 0 01.948.684l1.498 4.493a1 1 0 01-.502 1.21l-2.257 1.13a11.042 11.042 0 005.516 5.516l1.13-2.257a1 1 0 011.21-.502l4.493 1.498a1 1 0 01.684.949V19a2 2 0 01-2 2h-1C9.716 21 3 14.284 3 6V5z"
                  />
                </svg>
                {faculty.staffMobile}
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

export default FacultyCard;