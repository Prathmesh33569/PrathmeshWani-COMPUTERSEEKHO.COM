import React, { useState, useEffect } from 'react';
import { UserCheck } from 'lucide-react';
import RegisteredStudentList from './RegisteredStudentList';
import SearchFilter from './SearchFilter';
import { getAllRegisteredStudents } from '../../../Services/studentService';

const RegisteredStudentsPage = () => {
  const [students, setStudents] = useState([]);
  const [filteredStudents, setFilteredStudents] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null); // ✅ ADD ERROR STATE
  const [searchTerm, setSearchTerm] = useState('');
  const [filterCourse, setFilterCourse] = useState('');
  const [filterBatch, setFilterBatch] = useState('');

  useEffect(() => {
    fetchRegisteredStudents();
  }, []);

  useEffect(() => {
    applyFilters();
  }, [searchTerm, filterCourse, filterBatch, students]);

  const fetchRegisteredStudents = async () => {
    setLoading(true);
    setError(null); // ✅ RESET ERROR
    try {
      const data = await getAllRegisteredStudents();
      console.log('✅ Registered students from API:', data); // ✅ DEBUG LOG
      console.log('✅ First student:', data[0]); // ✅ DEBUG LOG
      
      // ✅ VALIDATE DATA
      if (!Array.isArray(data)) {
        throw new Error('API did not return an array');
      }
      
      setStudents(data);
      setFilteredStudents(data);
    } catch (error) {
      console.error('❌ Error fetching registered students:', error);
      setError(error.message); // ✅ SET ERROR
      alert('Error loading registered students: ' + error.message);
    } finally {
      setLoading(false);
    }
  };

  const applyFilters = () => {
    let filtered = [...students];

    // Search filter
    if (searchTerm) {
      filtered = filtered.filter(student =>
        student.studentName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
        student.studentMobile?.includes(searchTerm)
      );
    }

    // Course filter
    if (filterCourse) {
      filtered = filtered.filter(student => student.courseId === parseInt(filterCourse));
    }

    // Batch filter
    if (filterBatch) {
      filtered = filtered.filter(student => student.batchId === parseInt(filterBatch));
    }

    setFilteredStudents(filtered);
  };

  if (loading) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-center">
          <div className="animate-spin rounded-full h-16 w-16 border-b-2 border-purple-600 mx-auto"></div>
          <p className="mt-4 text-gray-600">Loading registered students...</p>
        </div>
      </div>
    );
  }

  // ✅ SHOW ERROR STATE
  if (error) {
    return (
      <div className="flex items-center justify-center min-h-screen">
        <div className="text-center bg-red-50 p-8 rounded-lg">
          <p className="text-red-600 font-semibold mb-4">Error loading students</p>
          <p className="text-gray-600 mb-4">{error}</p>
          <button
            onClick={fetchRegisteredStudents}
            className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="mb-6">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-3">
              <UserCheck size={32} className="text-purple-600" />
              <h1 className="text-3xl font-bold text-gray-800">Registered Students</h1>
            </div>
          </div>
        </div>

        {/* Search & Filter */}
        <SearchFilter
          searchTerm={searchTerm}
          setSearchTerm={setSearchTerm}
          filterCourse={filterCourse}
          setFilterCourse={setFilterCourse}
          filterBatch={filterBatch}
          setFilterBatch={setFilterBatch}
          students={students}
        />

        {/* Student List */}
        <RegisteredStudentList
          students={filteredStudents}
          loading={loading}
          onRefresh={fetchRegisteredStudents}
        />
      </div>
    </div>
  );
};

export default RegisteredStudentsPage;