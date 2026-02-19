// import React, { useState } from 'react';
// import BatchManagement from './BatchManagement';
// import CourseManagement from './CourseManagement';

// const TableMaintenance = () => {
//   const [selectedTable, setSelectedTable] = useState('');

//   const tables = [
//     { value: 'batch', label: 'Batch Master' },
//     { value: 'course', label: 'Course Master' }
//   ];

//   return (
//     <div className="min-h-screen bg-gray-50 p-6">
//       <div className="max-w-7xl mx-auto">
//         {/* Header */}
//         <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
//           <h1 className="text-3xl font-bold text-gray-900 mb-2">Table Maintenance</h1>
//           <p className="text-gray-600">Manage master tables - Create, Update, and Delete records</p>
//         </div>

//         {/* Table Selection */}
//         <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
//           <label className="block text-sm font-medium text-gray-700 mb-2">
//             Select Master Table
//           </label>
//           <select
//             value={selectedTable}
//             onChange={(e) => setSelectedTable(e.target.value)}
//             className="w-full md:w-96 px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
//           >
//             <option value="">-- Select a Table --</option>
//             {tables.map(table => (
//               <option key={table.value} value={table.value}>
//                 {table.label}
//               </option>
//             ))}
//           </select>
//         </div>

//         {/* Render Selected Component */}
//         {selectedTable === 'batch' && <BatchManagement />}
//         {selectedTable === 'course' && <CourseManagement />}
//       </div>
//     </div>
//   );
// };

// export default TableMaintenance;

import React, { useState } from 'react';
import BatchManagement from './BatchManagement';
import CourseManagement from './CourseManagement';
import AnnouncementManagement from './AnnouncementMaintainance';

const TableMaintenance = () => {
  const [selectedTable, setSelectedTable] = useState('');

  const tables = [
    { value: 'batch', label: 'Batch Master' },
    { value: 'course', label: 'Course Master' },
    { value: 'announcement', label: 'Announcement Master' }
  ];

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Table Maintenance</h1>
          <p className="text-gray-600">Manage master tables - Create, Update, and Delete records</p>
        </div>

        {/* Table Selection */}
        <div className="bg-white rounded-lg shadow-sm p-6 mb-6">
          <label className="block text-sm font-medium text-gray-700 mb-2">
            Select Master Table
          </label>
          <select
            value={selectedTable}
            onChange={(e) => setSelectedTable(e.target.value)}
            className="w-full md:w-96 px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          >
            <option value="">-- Select a Table --</option>
            {tables.map(table => (
              <option key={table.value} value={table.value}>
                {table.label}
              </option>
            ))}
          </select>
        </div>

        {/* Render Selected Component */}
        {selectedTable === 'batch' && <BatchManagement />}
        {selectedTable === 'course' && <CourseManagement />}
        {selectedTable === 'announcement' && <AnnouncementManagement />}
      </div>
    </div>
  );
};

export default TableMaintenance;