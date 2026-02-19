// import React from "react";
// import { useNavigate } from "react-router-dom";

// const Courses = ({ title, description, slug }) => {
//   const navigate = useNavigate();

//   const handleClick = () => {
//     navigate(`/courses/${slug}`);
//   };

//   return (
//     <div className="bg-white w-[330px] p-8 rounded-xl shadow-xl text-center">
//       <h3 className="text-xl font-semibold text-[#0b2c5f] mb-4">{title}</h3>

//       <p className="text-gray-600 leading-relaxed">{description}</p>

//       <button
//         onClick={handleClick}
//         className="mt-6 bg-red-600 text-white px-6 py-2 rounded-full hover:bg-red-700 transition"
//       >
//         Click to know more
//       </button>
//     </div>
//   );
// };

// export default Courses;

import React from "react";
import { useNavigate } from "react-router-dom";

const Courses = ({ title, description, slug }) => {
  const navigate = useNavigate(); // ✅ MUST be inside component

  const handleClick = () => {
    navigate(`/courses/${slug}`); // ✅ now navigate works
  };

  return (
    <div className="bg-white w-[330px] p-8 rounded-xl shadow-xl text-center">
      <h3 className="text-xl font-semibold text-[#0b2c5f] mb-4">{title}</h3>

      <p className="text-gray-600 leading-relaxed">{description}</p>

      <button
        onClick={handleClick}
        className="mt-6 bg-red-600 text-white px-6 py-2 rounded-full hover:bg-red-700 transition"
      >
        Click to know more
      </button>
    </div>
  );
};

export default Courses;
