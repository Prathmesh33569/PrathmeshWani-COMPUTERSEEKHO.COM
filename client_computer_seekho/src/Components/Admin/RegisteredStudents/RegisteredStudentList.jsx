import React from 'react';
import { Users } from 'lucide-react';
import RegisteredStudentCard from './RegisteredStudentCard';

const RegisteredStudentList = ({ students, loading, onRefresh }) => {
  if (loading) {
    return null;
  }

  if (students.length === 0) {
    return (
      <div className="bg-white rounded-lg shadow p-12 text-center">
        <Users size={48} className="mx-auto text-gray-400 mb-4" />
        <h3 className="text-xl font-semibold text-gray-700 mb-2">No Registered Students Found</h3>
        <p className="text-gray-500">No students match your search criteria.</p>
      </div>
    );
  }

  return (
    <div className="space-y-4">
      {/* Stats */}
      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-6">
        <div className="bg-white rounded-lg shadow p-4">
          <div className="text-sm text-gray-600">Total Students</div>
          <div className="text-2xl font-bold text-purple-600">{students.length}</div>
        </div>
        <div className="bg-white rounded-lg shadow p-4">
          <div className="text-sm text-gray-600">Total Fees Collected</div>
          <div className="text-2xl font-bold text-green-600">
            ₹{students.reduce((sum, s) => sum + (parseFloat(s.amount) || 0), 0).toFixed(2)}
          </div>
        </div>
        <div className="bg-white rounded-lg shadow p-4">
          <div className="text-sm text-gray-600">Receipts Generated</div>
          <div className="text-2xl font-bold text-blue-600">
            {students.filter(s => s.receiptId).length}
          </div>
        </div>
      </div>

      {/* Student Cards */}
      <div className="grid gap-4">
        {students.map((student) => (
          <RegisteredStudentCard
            key={student.studentId}
            student={student}
            onRefresh={onRefresh}
          />
        ))}
      </div>
    </div>
  );
};

export default RegisteredStudentList;