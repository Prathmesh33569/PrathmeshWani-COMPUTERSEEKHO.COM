import React, { useState, useEffect } from 'react';
import { X } from 'lucide-react';
import { ENQUIRY_SOURCES, LOGGED_IN_STAFF_ID } from '../../../utils/constants';
import { createEnquiry, getAllCourses } from '../../../Services/enquiryService';

const AddEnquiryModal = ({ onClose, onSuccess }) => {
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
  const [courses, setCourses] = useState([]);
  const [loadingCourses, setLoadingCourses] = useState(true);

  useEffect(() => {
    fetchCourses();
  }, []);

  const fetchCourses = async () => {
    setLoadingCourses(true);
    try {
      const data = await getAllCourses();
      setCourses(data);
    } catch (error) {
      console.error('Error fetching courses:', error);
      alert('Error loading courses. Please refresh the page.');
    } finally {
      setLoadingCourses(false);
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
        staffId: LOGGED_IN_STAFF_ID,
        followupDate: formData.followupDate || null,
        specialInstructions: formData.specialInstructions || null
      };

      const result = await createEnquiry(payload);
      alert('Enquiry added successfully! Enquiry ID: ' + result.enquiryId);
      onSuccess();
      onClose();
    } catch (error) {
      alert('Error creating enquiry: ' + (error.response?.data?.message || error.message));
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50">
      <div className="bg-white rounded-lg shadow-xl max-w-3xl w-full max-h-[90vh] overflow-y-auto">
        <div className="sticky top-0 bg-green-600 text-white p-4 flex justify-between items-center">
          <h2 className="text-xl font-bold">Add New Enquiry</h2>
          <button onClick={onClose} className="hover:bg-green-700 p-1 rounded">
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Enquiry Source *
              </label>
              <select
                value={formData.enquirySource}
                onChange={(e) => handleChange('enquirySource', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
              />
            </div>

            <div>
              <label className="block text-sm font-medium text-gray-700 mb-2">
                Follow-up Date (Optional, default +3 days)
              </label>
              <input
                type="date"
                value={formData.followupDate}
                onChange={(e) => handleChange('followupDate', e.target.value)}
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
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
                className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
              />
            </div>
          </div>

          <div className="flex gap-3 mt-6">
            <button
              onClick={handleSubmit}
              disabled={loading}
              className="flex-1 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
            >
              {loading ? 'Adding...' : 'Add Enquiry'}
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

export default AddEnquiryModal;