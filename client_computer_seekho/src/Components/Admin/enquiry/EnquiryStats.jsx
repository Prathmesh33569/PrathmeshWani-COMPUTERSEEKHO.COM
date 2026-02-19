import React from "react";
import { isPending, isToday } from "../../../utils/helpers";

const EnquiryStats = ({ followups }) => {
  const todayCount = followups.filter(
    (f) => isToday(f.followupDate) && !f.isClosed,
  ).length;
  const pendingCount = followups.filter(
    (f) => isPending(f.followupDate) && !f.isClosed,
  ).length;
  const totalOpen = followups.filter((f) => !f.isClosed).length;

  return (
    <div className="grid grid-cols-4 gap-4">
      <div className="bg-white rounded-lg shadow p-4">
        <div className="text-sm text-gray-600">Today's Follow-ups</div>
        <div className="text-2xl font-bold text-green-600">{todayCount}</div>
      </div>
      {/* <div className="bg-white rounded-lg shadow p-4">
        <div className="text-sm text-gray-600">Pending (Overdue)</div>
        <div className="text-2xl font-bold text-red-600">{pendingCount}</div>
      </div> */}
      <div className="bg-white rounded-lg shadow p-4">
        <div className="text-sm text-gray-600">Total Open Enquiries</div>
        <div className="text-2xl font-bold text-blue-600">{totalOpen}</div>
      </div>
      <div className="bg-white rounded-lg shadow p-4">
        <div className="text-sm text-gray-600">Total Enquiries</div>
        <div className="text-2xl font-bold text-gray-700">
          {followups.length}
        </div>
      </div>
    </div>
  );
};

export default EnquiryStats;
