import { NavLink } from "react-router-dom";
import { useEffect, useState } from "react";
import axios from "axios";

function CoursesDropdown() {
  const [courses, setCourses] = useState([]);

  useEffect(() => {
    axios
      .get("http://localhost:8080/api/courses")
      .then((res) => setCourses(res.data))
      .catch((err) => console.error(err));
  }, []);

  return (
    <li className="relative group list-none">
      <span className="cursor-pointer text-blue-900 font-semibold">
        COURSES
      </span>

      <ul className="absolute top-full left-0 w-64 bg-white shadow-lg rounded-md hidden group-hover:block z-50">
        {courses.map((course) => {
          const slug = course.courseName
            .toLowerCase()
            .replace(/\s+/g, "-");

          return (
            <li key={course.courseId}>
              <NavLink
                to={`/courses/${slug}`}
                className="block px-4 py-3 text-blue-900 font-semibold hover:bg-blue-50"
              >
                {course.courseName}
              </NavLink>
            </li>
          );
        })}
      </ul>
    </li>
  );
}

export default CoursesDropdown;
