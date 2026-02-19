import React from 'react';
import { Plus } from 'lucide-react';
import { VIEW_TYPES } from '../../../utils/constants';

const Navigation = ({ activeView, onViewChange, onAddEnquiry }) => {
  return (
    <div className="bg-white shadow">
      <div className="max-w-7xl mx-auto px-4">
        <div className="flex gap-4 py-3">
          <button
            onClick={() => onViewChange(VIEW_TYPES.MY_FOLLOWUPS)}
            className={`px-4 py-2 rounded-lg font-medium transition ${
              activeView === VIEW_TYPES.MY_FOLLOWUPS
                ? 'bg-blue-600 text-white'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
            }`}
          >
            My Follow-ups
          </button>
          <button
            onClick={() => onViewChange(VIEW_TYPES.ALL_FOLLOWUPS)}
            className={`px-4 py-2 rounded-lg font-medium transition ${
              activeView === VIEW_TYPES.ALL_FOLLOWUPS
                ? 'bg-blue-600 text-white'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
            }`}
          >
            View All
          </button>
          <button
            onClick={onAddEnquiry}
            className="ml-auto px-4 py-2 bg-green-600 text-white rounded-lg font-medium hover:bg-green-700 transition flex items-center gap-2"
          >
            <Plus size={20} />
            Add Enquiry
          </button>
        </div>
      </div>
    </div>
  );
};

export default Navigation;