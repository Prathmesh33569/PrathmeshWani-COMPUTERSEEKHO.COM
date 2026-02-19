import React, { useState, useEffect } from 'react';
import { X, DollarSign, CreditCard } from 'lucide-react';
import { createPayment } from '../../../Services/studentService';

const PaymentModal = ({ enquiryData, onClose, onPaymentSuccess }) => {
  const [formData, setFormData] = useState({
    enquiryId: enquiryData.enquiryId,
    courseId: enquiryData.courseId,
    batchId: enquiryData.batchId,
    paymentTypeId: '',
    amount: enquiryData.courseFees // From batch/course
  });
  const [loading, setLoading] = useState(false);
  const [paymentTypes, setPaymentTypes] = useState([
    { id: 1, name: 'Cheque' },
    { id: 2, name: 'DD' },
    { id: 3, name: 'Bank Transfer' }
  ]);

  const handleSubmit = async () => {
    // Validation
    if (!formData.paymentTypeId) {
      alert('Please select payment type');
      return;
    }

    setLoading(true);
    try {
      const payload = {
        enquiryId: formData.enquiryId,
        courseId: formData.courseId,
        batchId: formData.batchId,
        paymentTypeId: parseInt(formData.paymentTypeId),
        amount: parseFloat(formData.amount)
      };

      const paymentResponse = await createPayment(payload);
      alert(`Payment successful! Payment ID: ${paymentResponse.paymentId}, Receipt ID: ${paymentResponse.receiptId}`);
      onPaymentSuccess(paymentResponse.paymentId);
      onClose();
    } catch (error) {
      alert('Payment failed: ' + (error.response?.data || error.message));
      // Keep modal open on failure
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-md w-full">
        <div className="sticky top-0 bg-green-600 text-white p-4 flex justify-between items-center">
          <div className="flex items-center gap-2">
            <CreditCard size={24} />
            <h2 className="text-xl font-bold">Payment Details</h2>
          </div>
          <button onClick={onClose} className="hover:bg-green-700 p-1 rounded">
            <X size={24} />
          </button>
        </div>

        <div className="p-6">
          <div className="bg-blue-50 border-l-4 border-blue-600 p-4 mb-6">
            <p className="text-sm text-blue-800">
              <strong>Note:</strong> Complete payment to proceed with student registration.
            </p>
          </div>

          <div className="space-y-4">
            {/* Amount Display */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Course Fees *
              </label>
              <div className="flex items-center gap-2 px-3 py-2 border border-gray-300 rounded-lg bg-gray-50">
                <DollarSign size={20} className="text-green-600" />
                <span className="text-2xl font-bold text-gray-800">₹{formData.amount}</span>
              </div>
            </div>

            {/* Payment Type Selection */}
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Payment Method *
              </label>
              <select
                value={formData.paymentTypeId}
                onChange={(e) => setFormData({...formData, paymentTypeId: e.target.value})}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-green-500"
                required
              >
                <option value="">Select Payment Method</option>
                {paymentTypes.map(type => (
                  <option key={type.id} value={type.id}>
                    {type.name}
                  </option>
                ))}
              </select>
            </div>

            {/* Course & Batch Info */}
            <div className="bg-gray-50 rounded-lg p-4">
              <h3 className="font-semibold text-gray-800 mb-2">Payment For:</h3>
              <div className="text-sm text-gray-600 space-y-1">
                <div><span className="font-medium">Enquiry ID:</span> #{enquiryData.enquiryId}</div>
                <div><span className="font-medium">Course ID:</span> {enquiryData.courseId}</div>
                <div><span className="font-medium">Batch ID:</span> {enquiryData.batchId}</div>
              </div>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex gap-3 mt-6">
            <button
              onClick={handleSubmit}
              disabled={loading}
              className="flex-1 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {loading ? 'Processing Payment...' : 'Confirm Payment'}
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

export default PaymentModal;