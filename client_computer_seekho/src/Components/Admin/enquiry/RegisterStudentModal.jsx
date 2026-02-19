import React, { useState, useEffect } from 'react';
import { X, User } from 'lucide-react';
import { registerStudent, getBatchesByCourse } from '../../../Services/studentService';
import { getEnquiryById } from '../../../Services/enquiryService';
import PaymentModal from './PaymentModal';

const RegisterStudentModal = ({ enquiryId, onClose, onSuccess }) => {
  const [formData, setFormData] = useState({
    enquiryId: enquiryId,
    studentName: '',
    studentMobile: '',
    studentAddress: '',
    studentGender: '',
    studentDob: '',
    studentQualification: '',
    photoUrl: '',
    courseId: '',
    batchId: '',
    studentUsername: '',
    studentPassword: '',
    paymentId: null
  });
  
  const [loading, setLoading] = useState(false);
  const [fetching, setFetching] = useState(true);
  const [batches, setBatches] = useState([]);
  const [loadingBatches, setLoadingBatches] = useState(false);
  const [showPaymentModal, setShowPaymentModal] = useState(false);
  const [courseFees, setCourseFees] = useState(0);

  useEffect(() => {
    fetchEnquiryDetails();
  }, [enquiryId]);

  useEffect(() => {
    if (formData.courseId) {
      fetchBatchesForCourse(formData.courseId);
    }
  }, [formData.courseId]);

  const fetchEnquiryDetails = async () => {
    setFetching(true);
    try {
      const enquiry = await getEnquiryById(enquiryId);
      
      console.log('Fetched enquiry:', enquiry);
      
      // Pre-fill from enquiry
      setFormData(prev => ({
        ...prev,
        studentName: enquiry.enquirerName || enquiry.studentName || '',
        studentMobile: enquiry.enquirerMobile || '',
        studentAddress: enquiry.enquirerAddress || '',
        courseId: enquiry.course?.courseId?.toString() || ''
      }));
      
      console.log('CourseId set to:', enquiry.course?.courseId);
    } catch (error) {
      alert('Error fetching enquiry details: ' + (error.response?.data?.message || error.message));
      onClose();
    } finally {
      setFetching(false);
    }
  };

  const fetchBatchesForCourse = async (courseId) => {
    console.log('Fetching batches for courseId:', courseId);
    setLoadingBatches(true);
    try {
      const data = await getBatchesByCourse(courseId);
      console.log('Received batches:', data);
      setBatches(data);
      
      if (data.length === 0) {
        console.warn('No batches found for course:', courseId);
        alert('No active batches found for this course. Please create a batch first.');
      }
    } catch (error) {
      console.error('Error fetching batches:', error);
      setBatches([]);
      alert('Error loading batches: ' + (error.response?.data?.message || error.message));
    } finally {
      setLoadingBatches(false);
    }
  };

  const handleChange = (field, value) => {
    setFormData({ ...formData, [field]: value });
    
    // When batch is selected, get course fees
    if (field === 'batchId' && value) {
      const selectedBatch = batches.find(b => (b.batchId || b.batch_id) === parseInt(value));
      if (selectedBatch) {
        const fees = selectedBatch.courseFees || selectedBatch.course_fees || 0;
        setCourseFees(fees);
        console.log('Course fees set to:', fees);
      }
    }
  };

  const handleProceedToPayment = () => {
    // Validation before showing payment modal
    if (!formData.studentName || !formData.studentMobile || !formData.studentGender || 
        !formData.batchId || !formData.studentUsername || !formData.studentPassword) {
      alert('Please fill all required fields before proceeding to payment');
      return;
    }

    if (!courseFees || courseFees === 0) {
      alert('Course fees not found. Please select a valid batch.');
      return;
    }

    setShowPaymentModal(true);
  };

  const handlePaymentSuccess = async (paymentId) => {
    // After payment success, register student with paymentId
    setShowPaymentModal(false);
    setLoading(true);
    
    try {
      const payload = {
        enquiryId: formData.enquiryId,
        studentName: formData.studentName,
        studentMobile: parseInt(formData.studentMobile),
        studentAddress: formData.studentAddress || null,
        studentGender: formData.studentGender,
        studentDob: formData.studentDob || null,
        studentQualification: formData.studentQualification || null,
        photoUrl: formData.photoUrl || null,
        courseId: parseInt(formData.courseId),
        batchId: parseInt(formData.batchId),
        studentUsername: formData.studentUsername,
        studentPassword: formData.studentPassword,
        paymentId: paymentId
      };

      console.log('Registering student with payload:', payload);
      const result = await registerStudent(payload);
      
      alert('Student registered successfully! Student ID: ' + result.studentId);
      onSuccess();
      onClose();
    } catch (error) {
      alert('Error registering student: ' + (error.response?.data || error.message));
    } finally {
      setLoading(false);
    }
  };

  if (fetching) {
    return (
      <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
        <div className="bg-white rounded-lg shadow-xl p-8">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-purple-600 mx-auto"></div>
          <p className="mt-4 text-gray-600">Loading enquiry details...</p>
        </div>
      </div>
    );
  }

  return (
    <>
      <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
        <div className="bg-white rounded-lg shadow-xl max-w-4xl w-full max-h-[90vh] overflow-y-auto">
          <div className="sticky top-0 bg-purple-600 text-white p-4 flex justify-between items-center">
            <div className="flex items-center gap-2">
              <User size={24} />
              <h2 className="text-xl font-bold">Student Registration - Enquiry #{enquiryId}</h2>
            </div>
            <button onClick={onClose} className="hover:bg-purple-700 p-1 rounded" disabled={loading}>
              <X size={24} />
            </button>
          </div>

          <div className="p-6">
            <div className="bg-blue-50 border-l-4 border-blue-600 p-4 mb-6">
              <p className="text-sm text-blue-800">
                <strong>Note:</strong> Fill student details and proceed to payment. After payment confirmation, registration will be completed.
              </p>
            </div>

            <div className="grid grid-cols-2 gap-4">
              {/* Basic Details */}
              <div className="col-span-2">
                <h3 className="text-lg font-semibold text-gray-800 mb-3 pb-2 border-b">Basic Details</h3>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Student Name *
                </label>
                <input
                  type="text"
                  value={formData.studentName}
                  onChange={(e) => handleChange('studentName', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  required
                  disabled={loading}
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Mobile Number *
                </label>
                <input
                  type="tel"
                  value={formData.studentMobile}
                  onChange={(e) => handleChange('studentMobile', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  required
                  disabled={loading}
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Gender *
                </label>
                <select
                  value={formData.studentGender}
                  onChange={(e) => handleChange('studentGender', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  required
                  disabled={loading}
                >
                  <option value="">Select Gender</option>
                  <option value="Male">Male</option>
                  <option value="Female">Female</option>
                  <option value="Other">Other</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Date of Birth
                </label>
                <input
                  type="date"
                  value={formData.studentDob}
                  onChange={(e) => handleChange('studentDob', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  disabled={loading}
                />
              </div>

              <div className="col-span-2">
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Address
                </label>
                <textarea
                  value={formData.studentAddress}
                  onChange={(e) => handleChange('studentAddress', e.target.value)}
                  rows="2"
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  disabled={loading}
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Qualification
                </label>
                <select
                  value={formData.studentQualification}
                  onChange={(e) => handleChange('studentQualification', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  disabled={loading}
                >
                  <option value="">Select Qualification</option>
                  <option value="10th">10th Pass</option>
                  <option value="12th">12th Pass</option>
                  <option value="Graduate">Graduate</option>
                  <option value="Post-Graduate">Post-Graduate</option>
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Photo URL (Optional)
                </label>
                <input
                  type="text"
                  value={formData.photoUrl}
                  onChange={(e) => handleChange('photoUrl', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  placeholder="https://example.com/photo.jpg"
                  disabled={loading}
                />
              </div>

              {/* Course & Batch */}
              <div className="col-span-2 mt-4">
                <h3 className="text-lg font-semibold text-gray-800 mb-3 pb-2 border-b">Course & Batch Details</h3>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Course (Pre-filled from Enquiry)
                </label>
                <input
                  type="text"
                  value={`Course ID: ${formData.courseId}`}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg bg-gray-50"
                  disabled
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Batch *
                </label>
                <select
                  value={formData.batchId}
                  onChange={(e) => handleChange('batchId', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  required
                  disabled={loadingBatches || loading}
                >
                  <option value="">
                    {loadingBatches ? 'Loading batches...' : 'Select Batch'}
                  </option>
                  {batches && batches.length > 0 ? (
                    batches.map(batch => (
                      <option key={batch.batchId || batch.batch_id} value={batch.batchId || batch.batch_id}>
                        {batch.batchName || batch.batch_name} - ₹{batch.courseFees || batch.course_fees || 0}
                      </option>
                    ))
                  ) : (
                    <option disabled>No batches available</option>
                  )}
                </select>
                <p className="text-xs text-gray-500 mt-1">
                  Found {batches.length} batch(es)
                </p>
              </div>

              {/* Course Fees Display */}
              {courseFees > 0 && (
                <div className="col-span-2">
                  <div className="bg-green-50 border-l-4 border-green-600 p-4">
                    <p className="text-sm text-green-800">
                      <strong>Course Fees:</strong> ₹{courseFees}
                    </p>
                  </div>
                </div>
              )}

              {/* Login Credentials */}
              <div className="col-span-2 mt-4">
                <h3 className="text-lg font-semibold text-gray-800 mb-3 pb-2 border-b">Login Credentials</h3>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Username *
                </label>
                <input
                  type="text"
                  value={formData.studentUsername}
                  onChange={(e) => handleChange('studentUsername', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  placeholder="e.g., sachin2026"
                  required
                  disabled={loading}
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Password *
                </label>
                <input
                  type="password"
                  value={formData.studentPassword}
                  onChange={(e) => handleChange('studentPassword', e.target.value)}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500"
                  required
                  disabled={loading}
                />
              </div>
            </div>

            <div className="flex gap-3 mt-6">
              <button
                onClick={handleProceedToPayment}
                disabled={loading}
                className="flex-1 px-4 py-2 bg-purple-600 text-white rounded-lg hover:bg-purple-700 transition font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
              >
                {loading ? 'Processing...' : 'Proceed to Payment'}
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

      {/* Payment Modal */}
      {showPaymentModal && (
        <PaymentModal
          enquiryData={{
            enquiryId: formData.enquiryId,
            courseId: formData.courseId,
            batchId: formData.batchId,
            courseFees: courseFees
          }}
          onClose={() => setShowPaymentModal(false)}
          onPaymentSuccess={handlePaymentSuccess}
        />
      )}
    </>
  );
};

export default RegisterStudentModal;