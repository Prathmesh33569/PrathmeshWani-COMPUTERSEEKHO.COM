import { useParams } from "react-router-dom";
import { useEffect, useState } from "react";
import axios from "axios";
import {
  FaClock,
  FaRupeeSign,
  FaUsers,
  FaFileAlt,
  FaInfoCircle,
} from "react-icons/fa";
import { useNavigate } from "react-router-dom";

function CoursesDetails() {
  const { slug } = useParams();
  const [course, setCourse] = useState(null);
  const [loading, setLoading] = useState(true);

  const navigate = useNavigate();

  useEffect(() => {
    window.scrollTo({ top: 0, behavior: "smooth" });

    axios
      .get("http://localhost:8080/api/courses")
      .then((res) => {
        const selectedCourse = res.data.find(
          (c) => c.courseName.toLowerCase().replace(/\s+/g, "-") === slug,
        );
        setCourse(selectedCourse || null);
        if (selectedCourse) {
          console.log(selectedCourse);
        }
      })
      .catch((err) => console.error(err))
      .finally(() => setLoading(false));
  }, [slug]);

  if (loading) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <p className="text-lg font-medium text-gray-600">
          Loading course details...
        </p>
      </div>
    );
  }

  if (!course) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-gray-50">
        <p className="text-lg font-medium text-red-600">Course not found</p>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gradient-to-br from-blue-50 via-purple-50 to-pink-50 px-6 py-12">
      <div className="max-w-5xl mx-auto bg-white/70 backdrop-blur-2xl rounded-3xl shadow-2xl px-10 py-14 border border-white/40">
        {/* TITLE */}
        <h1 className="text-4xl font-extrabold text-center mb-10 bg-gradient-to-r from-blue-600 to-purple-600 bg-clip-text text-transparent">
          {course.courseName}
        </h1>

        {/* COVER PHOTO */}
        {course.coverPhoto && course.coverPhoto.trim() !== "" && (
          <div className="flex justify-center mb-10">
            <img
              src={`http://localhost:8080${course.coverPhoto}`}
              alt={course.courseName}
              className="w-full max-w-3xl h-64 object-cover rounded-2xl shadow-lg"
            />
          </div>
        )}

        {/* DESCRIPTION SECTION */}
        {course.courseDescription && course.courseDescription.trim() !== "" && (
          <div className="max-w-3xl mx-auto mb-10">
            <div className="bg-gradient-to-r from-indigo-50 to-blue-50 rounded-2xl p-8 shadow-md border border-indigo-100">
              <h2 className="text-2xl font-bold text-gray-800 mb-4 flex items-center gap-2">
                <FaInfoCircle className="text-indigo-600" />
                Course Description
              </h2>
              <p className="text-gray-700 text-lg leading-relaxed text-justify">
                {course.courseDescription}
              </p>
            </div>
          </div>
        )}

        {/* INFO GRID */}
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-6 max-w-3xl mx-auto mb-10">
          {/* Duration */}
          {course.courseDuration && (
            <div className="flex items-center justify-between bg-gray-50 rounded-xl px-6 py-4 shadow-sm border-l-4 border-blue-500 transition-all duration-300 hover:shadow-lg hover:-translate-y-1 hover:bg-white">
              <span className="flex items-center gap-2 font-semibold text-gray-700">
                <FaClock className="text-blue-600" /> Duration
              </span>
              <span className="text-blue-900 font-bold">
                {course.courseDuration}{" "}
                {course.courseDuration === 1 ? "Month" : "Months"}
              </span>
            </div>
          )}

          {/* Fees */}
          {course.courseFees && (
            <div className="flex items-center justify-between bg-gray-50 rounded-xl px-6 py-4 shadow-sm border-l-4 border-green-500 transition-all duration-300 hover:shadow-lg hover:-translate-y-1 hover:bg-white">
              <span className="flex items-center gap-2 font-semibold text-gray-700">
                <FaRupeeSign className="text-green-600" /> Fees
              </span>
              <span className="text-blue-900 font-bold">
                ₹ {course.courseFees.toLocaleString("en-IN")}
              </span>
            </div>
          )}

          {/* Age Group */}
          {course.ageGrpType && course.ageGrpType.trim() !== "" && (
            <div className="flex items-center justify-between bg-gray-50 rounded-xl px-6 py-4 shadow-sm border-l-4 border-purple-500 transition-all duration-300 hover:shadow-lg hover:-translate-y-1 hover:bg-white">
              <span className="flex items-center gap-2 font-semibold text-gray-700">
                <FaUsers className="text-purple-600" /> Age Group
              </span>
              <span className="text-gray-800 font-medium">
                {course.ageGrpType}
              </span>
            </div>
          )}
        </div>

        {/* SYLLABUS SECTION */}
        {course.courseSyllabus && course.courseSyllabus.trim() !== "" && (
          <div className="max-w-3xl mx-auto mb-10">
            <div className="bg-gradient-to-r from-blue-50 to-purple-50 rounded-2xl p-8 shadow-md border border-blue-100">
              <h2 className="text-2xl font-bold text-gray-800 mb-4 flex items-center gap-2">
                <FaFileAlt className="text-blue-600" />
                Course Syllabus
              </h2>
              <div className="text-gray-700 text-lg leading-relaxed whitespace-pre-wrap">
                {course.courseSyllabus}
              </div>
            </div>
          </div>
        )}

        {/* ENROLL BUTTON */}
        <div className="flex justify-center mt-12">
          <button
            onClick={() => navigate("/contact")}
            className="px-10 py-4 bg-gradient-to-r from-blue-600 to-purple-600 text-white rounded-xl font-semibold shadow-lg hover:shadow-xl hover:scale-105 transition-all duration-300"
          >
            Enroll Now
          </button>
        </div>
      </div>
    </div>
  );
}

export default CoursesDetails;
