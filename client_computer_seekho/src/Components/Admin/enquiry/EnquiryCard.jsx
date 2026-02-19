import React from 'react';
import { Phone, Calendar, User, BookOpen, Edit, UserPlus } from 'lucide-react';
import { getStatusClass, getStatusTextClass, getStatusLabel, formatDate } from '../../../utils/helpers';

const EnquiryCard = ({ enquiry, onCallClick, onEditClick, onRegisterClick }) => {
  return (
    <div
      className={`bg-white rounded-lg shadow hover:shadow-md transition p-4 border-l-4 ${getStatusClass(enquiry.followupDate)}`}
    >
      <div className="flex items-start justify-between">
        <div className="flex-1 grid grid-cols-4 gap-4">
          {/* Column 1: Basic Info */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <User size={16} className="text-gray-500" />
              <span className="font-semibold text-gray-800">{enquiry.enquirerName}</span>
            </div>
            <div className="flex items-center gap-2 text-sm text-gray-600">
              <Phone size={14} />
              <span>{enquiry.enquirerMobile}</span>
            </div>
          </div>

          {/* Column 2: Course Info */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <BookOpen size={16} className="text-blue-600" />
              <span className="font-medium text-gray-700">{enquiry.courseName}</span>
            </div>
            <div className="text-sm text-gray-500">
              ID: #{enquiry.enquiryId}
            </div>
          </div>

          {/* Column 3: Follow-up Date */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <Calendar size={16} className="text-gray-500" />
              <span className="font-medium">Follow-up Date</span>
            </div>
            <div className={`text-sm font-semibold ${getStatusTextClass(enquiry.followupDate)}`}>
              {formatDate(enquiry.followupDate)}
              {getStatusLabel(enquiry.followupDate)}
            </div>
          </div>

          {/* Column 4: Action Buttons */}
          <div className="flex items-center justify-end gap-2">
            {enquiry.isClosed ? (
              <span className="px-3 py-1 bg-gray-200 text-gray-700 rounded-full text-sm font-medium">
                Closed
              </span>
            ) : (
              <>
                <button
                  onClick={() => onRegisterClick(enquiry)}
                  className="px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition flex items-center gap-2 font-medium"
                >
                  <UserPlus size={18} />
                  Register
                </button>
                <button
                  onClick={() => onEditClick(enquiry)}
                  className="px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition flex items-center gap-2 font-medium"
                >
                  <Edit size={18} />
                  Edit
                </button>
                <button
                  onClick={() => onCallClick(enquiry)}
                  className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition flex items-center gap-2 font-medium"
                >
                  <Phone size={18} />
                  Call
                </button>
              </>
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

export default EnquiryCard;