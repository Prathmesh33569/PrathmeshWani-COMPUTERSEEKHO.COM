import React, { useState, useEffect } from 'react';
import { X } from 'lucide-react';
import { CLOSURE_REASONS } from '../../../utils/constants';
import { getDefaultNextDate, formatDate } from '../../../utils/helpers';
import { updateFollowup } from '../../../Services/enquiryService';

const CallModal = ({ enquiry, onClose, onSuccess }) => {
  const [formData, setFormData] = useState({
    remarks: '',
    specialInstructions: '',
    nextFollowupDate: getDefaultNextDate(),
    closeEnquiry: false,
    closureReasonId: null,
    closureReasonText: ''
  });
  const [loading, setLoading] = useState(false);

  const handleSubmit = async () => {
    setLoading(true);
    try {
      const payload = {
        enquiryId: enquiry.enquiryId,
        remarks: formData.remarks,
        specialInstructions: formData.specialInstructions,
        nextFollowupDate: formData.closeEnquiry ? null : formData.nextFollowupDate,
        closeEnquiry: formData.closeEnquiry,
        closureReasonId: formData.closeEnquiry ? formData.closureReasonId : null,
        closureReasonText: formData.closeEnquiry && formData.closureReasonId === 4 
          ? formData.closureReasonText 
          : null
      };

      await updateFollowup(payload);
      alert('Follow-up updated successfully!');
      onSuccess();
      onClose();
    } catch (error) {
      alert('Error updating follow-up: ' + (error.response?.data?.message || error.message));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-2xl w-full max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-blue-600 text-white p-4 flex justify-between items-center">
          <h2 className="text-xl font-bold">Follow-up Call - {enquiry.enquirerName}</h2>
          <button onClick={onClose} className="hover:bg-blue-700 p-1 rounded">
            <X size={24} />
          </button>
        </div>

        <div className="p-6">
          {/* Enquiry Details */}
          <div className="bg-gray-50 rounded-lg p-4 mb-6">
            <h3 className="font-semibold text-gray-800 mb-3">Enquiry Details</h3>
            <div className="grid grid-cols-2 gap-3 text-sm">
              <div><span className="font-medium">Enquiry ID:</span> #{enquiry.enquiryId}</div>
              <div><span className="font-medium">Phone:</span> {enquiry.enquirerMobile}</div>
              <div><span className="font-medium">Course:</span> {enquiry.courseName}</div>
              <div><span className="font-medium">Current Follow-up:</span> {formatDate(enquiry.followupDate)}</div>
            </div>
          </div>

          {/* Form Fields */}
          <div className="space-y-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Conversation Remarks
              </label>
              <textarea
                value={formData.remarks}
                onChange={(e) => setFormData({...formData, remarks: e.target.value})}
                rows="3"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="Enter conversation details..."
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Special Instructions
              </label>
              <textarea
                value={formData.specialInstructions}
                onChange={(e) => setFormData({...formData, specialInstructions: e.target.value})}
                rows="2"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                placeholder="Any special notes or instructions..."
              />
            </div>

            <div>
              <label className="flex items-center gap-2 mb-3">
                <input
                  type="checkbox"
                  checked={formData.closeEnquiry}
                  onChange={(e) => setFormData({...formData, closeEnquiry: e.target.checked})}
                  className="w-4 h-4 text-blue-600 rounded focus:ring-2 focus:ring-blue-500"
                />
                <span className="text-sm font-medium text-gray-700">Close this enquiry</span>
              </label>

              {formData.closeEnquiry ? (
                <div className="space-y-3 ml-6">
                  <div>
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                      Closure Reason *
                    </label>
                    <select
                      value={formData.closureReasonId || ''}
                      onChange={(e) => setFormData({...formData, closureReasonId: parseInt(e.target.value)})}
                      className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                      required
                    >
                      <option value="">Select reason...</option>
                      {CLOSURE_REASONS.map(reason => (
                        <option key={reason.id} value={reason.id}>
                          {reason.label}
                        </option>
                      ))}
                    </select>
                  </div>

                  {formData.closureReasonId === 4 && (
                    <div>
                      <input
                        type="text"
                        value={formData.closureReasonText}
                        onChange={(e) => setFormData({...formData, closureReasonText: e.target.value})}
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                        placeholder="Specify other reason..."
                        required
                      />
                    </div>
                  )}
                </div>
              ) : (
                <div>
                  <label className="block text-sm font-medium text-gray-700 mb-2">
                    Next Follow-up Date (Default: +3 days)
                  </label>
                  <input
                    type="date"
                    value={formData.nextFollowupDate}
                    onChange={(e) => setFormData({...formData, nextFollowupDate: e.target.value})}
                    className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
                    required
                  />
                </div>
              )}
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex gap-3 mt-6">
            <button
              onClick={handleSubmit}
              disabled={loading}
              className="flex-1 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {loading ? 'Saving...' : 'Save Follow-up'}
            </button>
            <button
              onClick={onClose}
              disabled={loading}
              className="px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition font-medium disabled:bg-gray-100"
            >
              Cancel
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default CallModal;