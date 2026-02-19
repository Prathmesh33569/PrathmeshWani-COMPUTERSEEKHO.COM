import React, { useMemo } from 'react';
import { Search } from 'lucide-react';

const SearchFilter = ({ 
  searchTerm, 
  setSearchTerm, 
  filterCourse, 
  setFilterCourse,
  filterBatch,
  setFilterBatch,
  students 
}) => {
  
  // Get unique courses
  const courses = useMemo(() => {
    const uniqueCourses = [...new Set(students.map(s => s.courseId))];
    return uniqueCourses.filter(Boolean);
  }, [students]);

  // Get unique batches
  const batches = useMemo(() => {
    const uniqueBatches = [...new Set(students.map(s => s.batchId))];
    return uniqueBatches.filter(Boolean);
  }, [students]);

  return (
    <div className="bg-white rounded-lg shadow p-4 mb-6">
      <div className="flex gap-4 items-center">
        {/* Search */}
        <div className="flex-1 relative">
          <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400" size={20} />
          <input
            type="text"
            placeholder="Search by name or phone..."
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent"
          />
        </div>

        {/* Course Filter */}
        <select
          value={filterCourse}
          onChange={(e) => setFilterCourse(e.target.value)}
          className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
        >
          <option value="">All Courses</option>
          {courses.map(courseId => (
            <option key={courseId} value={courseId}>
              Course ID: {courseId}
            </option>
          ))}
        </select>

        {/* Batch Filter */}
        <select
          value={filterBatch}
          onChange={(e) => setFilterBatch(e.target.value)}
          className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
        >
          <option value="">All Batches</option>
          {batches.map(batchId => (
            <option key={batchId} value={batchId}>
              Batch ID: {batchId}
            </option>
          ))}
        </select>
      </div>
    </div>
  );
};

export default SearchFilter;