import React from 'react';
import { Search } from 'lucide-react';
import { FILTER_OPTIONS } from '../../../utils/constants';

const SearchFilter = ({ searchTerm, onSearchChange, filterStatus, onFilterChange }) => {
  return (
    <div className="bg-white rounded-lg shadow p-4">
      <div className="flex gap-4">
        <div className="flex-1 relative">
          <Search className="absolute left-3 top-3 text-gray-400" size={20} />
          <input
            type="text"
            placeholder="Search by name or phone..."
            value={searchTerm}
            onChange={(e) => onSearchChange(e.target.value)}
            className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <select
          value={filterStatus}
          onChange={(e) => onFilterChange(e.target.value)}
          className="px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
        >
          <option value={FILTER_OPTIONS.ALL}>All Follow-ups</option>
          <option value={FILTER_OPTIONS.TODAY}>Today</option>
          <option value={FILTER_OPTIONS.PENDING}>Pending (Overdue)</option>
        </select>
      </div>
    </div>
  );
};

export default SearchFilter;