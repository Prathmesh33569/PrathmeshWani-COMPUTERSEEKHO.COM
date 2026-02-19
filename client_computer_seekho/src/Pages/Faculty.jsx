// import React, { useEffect, useState } from "react";
// import { getFacultyList } from "../Services/facultyService";
// import FacultyCard from "../Components/Faculty/FacultyCard";

// function Faculty() {
//   const [facultyList, setFacultyList] = useState([]);
//   const [loading, setLoading] = useState(true);

//   useEffect(() => {
//     const fetchFaculty = async () => {
//       try {
//         const data = await getFacultyList();
//         setFacultyList(data);
//       } catch (error) {
//         console.error("Failed to load faculty", error);
//       } finally {
//         setLoading(false);
//       }
//     };
//     fetchFaculty();
//   }, []);

//   // if (loading) {
//   //   return (
//   //     <div className="py-20 text-center text-gray-500">
//   //       Loading faculty...
//   //     </div>
//   //   );
//   // }
//   if (loading) {
//   return (
//     <section className="py-12 bg-gray-50">
//       <div className="max-w-7xl mx-auto px-4">
//         <div className="text-center mb-10">
//           <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
//             Our Faculty
//           </h2>
//           <p className="text-gray-600 mt-2">
//             Learn from industry experts at Computer Seekho
//           </p>
//         </div>
//         <div className="flex flex-col gap-16">
//           {[1, 2, 3].map((i) => (
//             <div key={i} className="bg-white rounded-2xl shadow-md overflow-hidden animate-pulse">
//               <div className="w-full h-137.5 bg-gray-300"></div>
//               <div className="p-6">
//                 <div className="h-6 bg-gray-300 rounded w-1/3 mb-2"></div>
//                 <div className="h-4 bg-gray-300 rounded w-1/4 mb-4"></div>
//                 <div className="h-4 bg-gray-300 rounded w-full mb-2"></div>
//                 <div className="h-4 bg-gray-300 rounded w-5/6"></div>
//               </div>
//             </div>
//           ))}
//         </div>
//       </div>
//     </section>
//   );
// }

//   return (
//     <section className="py-12 bg-gray-50">
//       <div className="max-w-7xl mx-auto px-4">
//         {/* Heading */}
//         <div className="text-center mb-10">
//           <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
//             Our Faculty
//           </h2>
//           <p className="text-gray-600 mt-2">
//             Learn from industry experts at Computer Seekho
//           </p>
//         </div>

//         {/* ONE BELOW ANOTHER */}
//         <div className="flex flex-col gap-16">
//           {facultyList.map((faculty) => (
//             <FacultyCard key={faculty.id} faculty={faculty} />
//           ))}
//         </div>
//       </div>
//     </section>
//   );
// }

// export default Faculty;

import React, { useEffect, useState } from "react";
import { getFacultyList } from "../Services/facultyService";
import FacultyCard from "../Components/Faculty/FacultyCard";

function Faculty() {
  const [facultyList, setFacultyList] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    const fetchFaculty = async () => {
      try {
        setLoading(true);
        setError(null);
        const data = await getFacultyList();
        setFacultyList(data);
        setError(null);
      } catch (error) {
        console.error("Failed to load faculty", error);
        setError("Failed to load faculty members. Please try again later.");
      } finally {
        setLoading(false);
      }
    };
    fetchFaculty();
  }, []);

  // 🔄 LOADING STATE
  if (loading) {
    return (
      <section className="py-12 bg-gray-50 min-h-screen">
        <div className="max-w-7xl mx-auto px-4">
          <div className="text-center mb-10">
            <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
              Our Faculty
            </h2>
            <p className="text-gray-600 mt-2">
              Learn from industry experts at Computer Seekho
            </p>
          </div>

          {/* Loading Skeleton */}
          <div className="flex flex-col gap-8">
            {[1, 2, 3].map((i) => (
              <div
                key={i}
                className="bg-white rounded-xl shadow-lg overflow-hidden animate-pulse"
              >
                <div className="md:flex">
                  <div className="md:w-1/3 bg-gray-200 h-64"></div>
                  <div className="md:w-2/3 p-8 space-y-4">
                    <div className="h-8 bg-gray-200 rounded w-2/3"></div>
                    <div className="h-4 bg-gray-200 rounded w-1/2"></div>
                    <div className="h-20 bg-gray-200 rounded"></div>
                    <div className="flex gap-4">
                      <div className="h-4 bg-gray-200 rounded w-1/3"></div>
                      <div className="h-4 bg-gray-200 rounded w-1/4"></div>
                    </div>
                  </div>
                </div>
              </div>
            ))}
          </div>
        </div>
      </section>
    );
  }

  // ❌ ERROR STATE
  if (error) {
    return (
      <section className="py-12 bg-gray-50 min-h-screen">
        <div className="max-w-7xl mx-auto px-4">
          <div className="text-center mb-10">
            <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
              Our Faculty
            </h2>
            <p className="text-gray-600 mt-2">
              Learn from industry experts at Computer Seekho
            </p>
          </div>

          {/* Error Message */}
          <div className="max-w-md mx-auto">
            <div className="bg-red-50 border-l-4 border-red-500 p-6 rounded-lg">
              <div className="flex items-center mb-3">
                <svg
                  className="w-6 h-6 text-red-500 mr-3"
                  fill="none"
                  stroke="currentColor"
                  viewBox="0 0 24 24"
                >
                  <path
                    strokeLinecap="round"
                    strokeLinejoin="round"
                    strokeWidth="2"
                    d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"
                  />
                </svg>
                <h3 className="text-red-800 font-semibold">Error</h3>
              </div>
              <p className="text-red-700 text-sm">{error}</p>
              <button
                onClick={() => window.location.reload()}
                className="mt-4 px-4 py-2 bg-red-600 text-white rounded-lg hover:bg-red-700 transition text-sm font-semibold"
              >
                Retry
              </button>
            </div>
          </div>
        </div>
      </section>
    );
  }

  // 📭 EMPTY STATE
  if (facultyList.length === 0) {
    return (
      <section className="py-12 bg-gray-50 min-h-screen">
        <div className="max-w-7xl mx-auto px-4">
          <div className="text-center mb-10">
            <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
              Our Faculty
            </h2>
            <p className="text-gray-600 mt-2">
              Learn from industry experts at Computer Seekho
            </p>
          </div>

          {/* Empty State */}
          <div className="max-w-md mx-auto text-center py-12">
            <svg
              className="w-24 h-24 mx-auto text-gray-400 mb-6"
              fill="none"
              stroke="currentColor"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                strokeWidth="2"
                d="M17 20h5v-2a3 3 0 00-5.356-1.857M17 20H7m10 0v-2c0-.656-.126-1.283-.356-1.857M7 20H2v-2a3 3 0 015.356-1.857M7 20v-2c0-.656.126-1.283.356-1.857m0 0a5.002 5.002 0 019.288 0M15 7a3 3 0 11-6 0 3 3 0 016 0zm6 3a2 2 0 11-4 0 2 2 0 014 0zM7 10a2 2 0 11-4 0 2 2 0 014 0z"
              />
            </svg>
            <h3 className="text-xl font-semibold text-gray-800 mb-2">
              No Faculty Members Yet
            </h3>
            <p className="text-gray-600">
              Faculty members will appear here once they register.
            </p>
          </div>
        </div>
      </section>
    );
  }

  // ✅ SUCCESS STATE - DISPLAY FACULTY
  return (
    <section className="py-12 bg-gray-50 min-h-screen">
      <div className="max-w-7xl mx-auto px-4">
        {/* Heading */}
        <div className="text-center mb-10">
          <h2 className="text-3xl md:text-4xl font-bold text-gray-800">
            Our Faculty
          </h2>
          <p className="text-gray-600 mt-2">
            Learn from industry experts at Computer Seekho
          </p>
          <div className="mt-4">
            <span className="inline-block px-4 py-2 bg-purple-100 text-purple-700 rounded-full text-sm font-semibold">
              {facultyList.length} Faculty Member{facultyList.length !== 1 ? "s" : ""}
            </span>
          </div>
        </div>

        {/* Faculty Cards - ONE BELOW ANOTHER */}
        <div className="flex flex-col gap-8">
          {facultyList.map((faculty) => (
            <FacultyCard 
              key={faculty.staffId} 
              faculty={faculty} 
            />
          ))}
        </div>
      </div>
    </section>
  );
}

export default Faculty;