// import React, { useEffect, useState } from "react";
// import axios from "axios";
// import Courses from "./Courses";

// function CourseOffered() {
//   const [courses, setCourses] = useState([]);
//   const [loading, setLoading] = useState(true);

//   useEffect(() => {
//     axios
//       .get("http://localhost:8080/api/courses/active")
//       .then((res) => {
//         setCourses(res.data);
//       })
//       .catch((err) => {
//         console.error("Error fetching courses:", err);
//       })
//       .finally(() => {
//         setLoading(false);
//       });
//   }, []);

//   return (
//     <div className="bg-gray-50 py-16">
//       <h2 className="text-4xl font-bold text-center mb-12">Courses Offered</h2>

//       <div className="flex justify-center gap-10 flex-wrap">
//         {loading ? (
//           <p className="text-lg">Loading courses...</p>
//         ) : courses.length > 0 ? (
//           courses.map((course) => (
//             <Courses
//               key={course.courseId}
//               title={course.courseName}
//               description={course.courseDescription}
//             />
//           ))
//         ) : (
//           <p className="text-lg">No courses available</p>
//         )}
//       </div>
//     </div>
//   );
// }

// export default CourseOffered;




// import React, { useEffect, useState } from "react";
// import axios from "axios";
// import Courses from "./Courses";

// function CourseOffered() {
//   const [courses, setCourses] = useState([]);
//   const [loading, setLoading] = useState(true);

//   useEffect(() => {
//     axios
//       .get("http://localhost:8080/api/courses/active")
//       .then((res) => {
//         setCourses(res.data);
//       })
//       .catch((err) => {
//         console.error("Error fetching courses:", err);
//       })
//       .finally(() => {
//         setLoading(false);
//       });
//   }, []);

//   return (
//     <div className="bg-gray-50 py-16">
//       <h2 className="text-4xl font-bold text-center mb-12">Courses Offered</h2>

//       <div className="flex justify-center gap-10 flex-wrap">
//         {loading ? (
//           <p className="text-lg">Loading courses...</p>
//         ) : courses.length > 0 ? (
//           courses.map((course) => {
//             // ✅ CREATE SLUG HERE
//             const slug = course.courseName
//               .toLowerCase()
//               .replace(/\s+/g, "-");

//             return (
//               <Courses
//                 key={course.courseId}
//                 title={course.courseName}
//                 description={course.courseDescription}
//                 slug={slug} // ✅ PASS SLUG
//               />
//             );
//           })
//         ) : (
//           <p className="text-lg">No courses available</p>
//         )}
//       </div>
//     </div>
//   );
// }

// export default CourseOffered;


import React, { useEffect, useState } from "react";
import axios from "axios";
import Courses from "./Courses";

function CourseOffered() {
  const [courses, setCourses] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    axios
      .get("http://localhost:8080/api/courses/active")
      .then((res) => {
        setCourses(res.data);
      })
      .catch((err) => {
        console.error("Error fetching courses:", err);
      })
      .finally(() => {
        setLoading(false);
      });
  }, []);

  return (
    <div className="bg-gray-50 py-16">
      <h2 className="text-4xl font-bold text-center mb-12">Courses Offered</h2>

      <div className="flex justify-center gap-10 flex-wrap">
        {loading ? (
          <p className="text-lg">Loading courses...</p>
        ) : courses.length > 0 ? (
          courses.map((course) => {
            // ✅ CREATE SLUG HERE
            const slug = course.courseName
              .toLowerCase()
              .replace(/\s+/g, "-");

            return (
              <Courses
                key={course.courseId}
                title={course.courseName}
                description={course.courseDescription}
                slug={slug}
                coverPhoto={course.coverPhoto} // ✅ PASS COVER PHOTO
              />
            );
          })
        ) : (
          <p className="text-lg">No courses available</p>
        )}
      </div>
    </div>
  );
}

export default CourseOffered;
