import React, { useState, useEffect } from 'react';
import { X } from 'lucide-react';
import { ENQUIRY_SOURCES } from '../../../utils/constants';
import { getEnquiryById, updateEnquiry, getAllCourses } from '../../../Services/enquiryService';

const EditEnquiryModal = ({ enquiryId, onClose, onSuccess }) => {
  const [formData, setFormData] = useState({
    enquirerName: '',
    studentName: '',
    enquirerAddress: '',
    enquirerMobile: '',
    enquirerAlternateMobile: '',
    enquirerEmailId: '',
    enquirerQuery: '',
    enquirySource: 'Telephonic',
    courseId: '',
    followupDate: '',
    specialInstructions: ''
  });
  const [loading, setLoading] = useState(false);
  const [fetching, setFetching] = useState(true);
  const [courses, setCourses] = useState([]);
  const [loadingCourses, setLoadingCourses] = useState(true);

  useEffect(() => {
    fetchCourses();
    fetchEnquiryDetails();
  }, [enquiryId]);

  const fetchCourses = async () => {
    setLoadingCourses(true);
    try {
      const data = await getAllCourses();
      setCourses(data);
    } catch (error) {
      console.error('Error fetching courses:', error);
    } finally {
      setLoadingCourses(false);
    }
  };

  const fetchEnquiryDetails = async () => {
    setFetching(true);
    try {
      const enquiry = await getEnquiryById(enquiryId);
      
      setFormData({
        enquirerName: enquiry.enquirerName || '',
        studentName: enquiry.studentName || '',
        enquirerAddress: enquiry.enquirerAddress || '',
        enquirerMobile: enquiry.enquirerMobile?.toString() || '',
        enquirerAlternateMobile: enquiry.enquirerAlternateMobile?.toString() || '',
        enquirerEmailId: enquiry.enquirerEmailId || '',
        enquirerQuery: enquiry.enquirerQuery || '',
        enquirySource: enquiry.enquirySource || 'Telephonic',
        courseId: enquiry.course?.courseId?.toString() || '',
        followupDate: enquiry.followupDate || '',
        specialInstructions: enquiry.specialInstructions || ''
      });
    } catch (error) {
      alert('Error fetching enquiry details: ' + (error.response?.data?.message || error.message));
      onClose();
    } finally {
      setFetching(false);
    }
  };

  const handleChange = (field, value) => {
    setFormData({ ...formData, [field]: value });
  };

  const handleSubmit = async () => {
    // Validation
    if (!formData.enquirerName || !formData.enquirerMobile || !formData.courseId) {
      alert('Please fill all required fields (Enquirer Name, Mobile, Course ID)');
      return;
    }

    setLoading(true);
    try {
      const payload = {
        enquirerName: formData.enquirerName,
        studentName: formData.studentName || null,
        enquirerAddress: formData.enquirerAddress || null,
        enquirerMobile: parseInt(formData.enquirerMobile),
        enquirerAlternateMobile: formData.enquirerAlternateMobile 
          ? parseInt(formData.enquirerAlternateMobile) 
          : null,
        enquirerEmailId: formData.enquirerEmailId || null,
        enquirerQuery: formData.enquirerQuery || null,
        enquirySource: formData.enquirySource,
        courseId: parseInt(formData.courseId),
        followupDate: formData.followupDate || null,
        specialInstructions: formData.specialInstructions || null
      };

      await updateEnquiry(enquiryId, payload);
      alert('Enquiry updated successfully!');
      onSuccess();
      onClose();
    } catch (error) {
      alert('Error updating enquiry: ' + (error.response?.data?.message || error.message));
    } finally {
      setLoading(false);
    }
  };

  if (fetching) {
    return (
      <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
        <div className="bg-white rounded-lg shadow-xl p-8">
          <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600 mx-auto"></div>
          <p className="mt-4 text-gray-600">Loading enquiry details...</p>
        </div>
      </div>
    );
  }

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-orange-600 text-white p-4 flex justify-between items-center">
          <h2 className="text-xl font-bold">Edit Enquiry - #{enquiryId}</h2>
          <button onClick={onClose} className="hover:bg-orange-700 p-1 rounded">
            <X size={24} />
          </button>
        </div>

        <div className="p-6">
          <div className="grid grid-cols-2 gap-4">
            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Enquirer Name *
              </label>
              <input
                type="text"
                value={formData.enquirerName}
                onChange={(e) => handleChange('enquirerName', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
                required
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Student Name
              </label>
              <input
                type="text"
                value={formData.studentName}
                onChange={(e) => handleChange('studentName', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Mobile Number *
              </label>
              <input
                type="tel"
                value={formData.enquirerMobile}
                onChange={(e) => handleChange('enquirerMobile', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
                required
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Alternate Mobile
              </label>
              <input
                type="tel"
                value={formData.enquirerAlternateMobile}
                onChange={(e) => handleChange('enquirerAlternateMobile', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div className="col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Email
              </label>
              <input
                type="email"
                value={formData.enquirerEmailId}
                onChange={(e) => handleChange('enquirerEmailId', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div className="col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Address
              </label>
              <textarea
                value={formData.enquirerAddress}
                onChange={(e) => handleChange('enquirerAddress', e.target.value)}
                rows="2"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Enquiry Source *
              </label>
              <select
                value={formData.enquirySource}
                onChange={(e) => handleChange('enquirySource', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
                required
              >
                {ENQUIRY_SOURCES.map(source => (
                  <option key={source} value={source}>{source}</option>
                ))}
              </select>
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Course *
              </label>
              <select
                value={formData.courseId}
                onChange={(e) => handleChange('courseId', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
                required
                disabled={loadingCourses}
              >
                <option value="">
                  {loadingCourses ? 'Loading courses...' : 'Select a course'}
                </option>
                {courses.map(course => (
                  <option key={course.courseId} value={course.courseId}>
                    {course.courseName}
                  </option>
                ))}
              </select>
            </div>

            <div className="col-span-2">
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Query/Requirements
              </label>
              <textarea
                value={formData.enquirerQuery}
                onChange={(e) => handleChange('enquirerQuery', e.target.value)}
                rows="3"
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Follow-up Date
              </label>
              <input
                type="date"
                value={formData.followupDate}
                onChange={(e) => handleChange('followupDate', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Special Instructions
              </label>
              <input
                type="text"
                value={formData.specialInstructions}
                onChange={(e) => handleChange('specialInstructions', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-orange-500"
              />
            </div>
          </div>

          <div className="flex gap-3 mt-6">
            <button
              onClick={handleSubmit}
              disabled={loading}
              className="flex-1 px-4 py-2 bg-orange-600 text-white rounded-lg hover:bg-orange-700 transition font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {loading ? 'Updating...' : 'Update Enquiry'}
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

export default EditEnquiryModal;