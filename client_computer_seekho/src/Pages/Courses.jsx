import React from "react";
import { useNavigate } from "react-router-dom";

const Courses = ({ title, description, slug, coverPhoto }) => {
  const navigate = useNavigate();

  const handleClick = () => {
    navigate(`/courses/${slug}`);
  };

  return (
    <div className="bg-white w-[330px] rounded-xl shadow-xl overflow-hidden hover:shadow-2xl transition-shadow duration-300 flex flex-col">
      {/* Course Image */}
      <div className="h-48 w-full bg-gray-200">
        {coverPhoto ? (
          <img
            src={coverPhoto}
            alt={title}
            className="w-full h-full object-cover"
            onError={(e) => {
              e.target.src = 'https://via.placeholder.com/330x200?text=Course+Image';
            }}
          />
        ) : (
          <div className="w-full h-full flex items-center justify-center text-gray-400">
            No Image
          </div>
        )}
      </div>

      <div className="p-6 flex flex-col flex-1">
        <h3 className="text-xl font-semibold text-[#0b2c5f] mb-3">{title}</h3>

        <p className="text-gray-600 text-sm leading-relaxed mb-6 line-clamp-3 flex-1">
          {description}
        </p>

        <button
          onClick={handleClick}
          className="w-full mt-auto bg-red-600 text-white px-6 py-2.5 rounded-full hover:bg-red-700 transition font-medium"
        >
          View Details
        </button>
      </div>
    </div>
  );
};

export default Courses;
