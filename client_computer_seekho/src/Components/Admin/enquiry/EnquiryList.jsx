import React from 'react';
import { Calendar } from 'lucide-react';
import EnquiryCard from './EnquiryCard';

const EnquiryList = ({ enquiries, onCallClick, onEditClick, onRegisterClick, loading }) => {
  if (loading) {
    return null; // Loading is handled by parent
  }

  if (enquiries.length === 0) {
    return (
      <div className="bg-white rounded-lg shadow p-12 text-center">
        <Calendar size={48} className="mx-auto text-gray-400 mb-4" />
        <h3 className="text-xl font-semibold text-gray-700 mb-2">No Follow-ups Found</h3>
        <p className="text-gray-500">There are no follow-ups matching your criteria.</p>
      </div>
    );
  }

  return (
    <div className="grid gap-4">
      {enquiries.map((enquiry) => (
        <EnquiryCard
          key={enquiry.enquiryId}
          enquiry={enquiry}
          onCallClick={onCallClick}
          onEditClick={onEditClick}
          onRegisterClick={onRegisterClick}
        />
      ))}
    </div>
  );
};

export default EnquiryList;