import React, { useState, useEffect } from "react";
import {
  Search,
  Plus,
  Edit2,
  Trash2,
  X,
  Save,
  Upload,
  Image as ImageIcon,
} from "lucide-react";
import Tablemaintainanceapi from "../../../Services/Tablemaintainanceapi";
import batchService from "../../../Services/batchService";
import Pagination from "./Pagination";

const BatchManagement = () => {
  const [batches, setBatches] = useState([]);
  const [activeCourses, setActiveCourses] = useState([]);
  const [showModal, setShowModal] = useState(false);
  const [modalMode, setModalMode] = useState("add");
  const [searchTerm, setSearchTerm] = useState("");
  const [loading, setLoading] = useState(false);
  const [currentBatch, setCurrentBatch] = useState(null);
  const [currentPage, setCurrentPage] = useState(1);
  const [itemsPerPage] = useState(3);
  const [batchLogoFile, setBatchLogoFile] = useState(null);

  useEffect(() => {
    fetchBatches();
    fetchActiveCourses();
  }, []);

  const fetchBatches = async () => {
    setLoading(true);
    try {
      const response = await Tablemaintainanceapi.get("/batches");
      console.log("Batches response:", response.data);
      setBatches(response.data);
    } catch (error) {
      console.error("Error fetching batches:", error);
      alert("Error fetching batches: " + error.message);
    } finally {
      setLoading(false);
    }
  };

  const fetchActiveCourses = async () => {
    try {
      const response = await Tablemaintainanceapi.get("/courses/active");
      console.log("Active courses response:", response.data);
      if (Array.isArray(response.data)) {
        setActiveCourses(response.data);
      } else {
        console.error("Courses data is not an array:", response.data);
        setActiveCourses([]);
      }
    } catch (error) {
      console.error("Error fetching courses:", error);
      alert(
        "Error fetching courses: " +
          (error.response?.data?.message || error.message),
      );
    }
  };

  const handleAdd = () => {
    setModalMode("add");
    setCurrentBatch({
      batchName: "",
      batchStartDate: "",
      batchEndDate: "",
      courseId: "",
      presentationDate: "",
      finalPresentationDate: "",
      courseFees: "",
      courseFeesFrom: "",
      courseFeesTo: "",
      batchIsActive: true,
      batchLogoUrl: "",
    });
    setBatchLogoFile(null);
    setShowModal(true);
  };

  const handleEdit = (batch) => {
    setModalMode("edit");
    setCurrentBatch({
      batchId: batch.batchId || "",
      batchName: batch.batchName || "",
      batchStartDate: batch.batchStartDate || "",
      batchEndDate: batch.batchEndDate || "",
      courseId: batch.courseId || "",
      presentationDate: batch.presentationDate
        ? batch.presentationDate.replace(" ", "T")
        : "",
      finalPresentationDate: batch.finalPresentationDate
        ? batch.finalPresentationDate.replace(" ", "T")
        : "",
      courseFees: batch.courseFees || "",
      courseFeesFrom: batch.courseFeesFrom || "",
      courseFeesTo: batch.courseFeesTo || "",
      batchIsActive:
        batch.batchIsActive !== undefined ? batch.batchIsActive : true,
      batchLogoUrl: batch.batchLogoUrl || "",
    });
    setBatchLogoFile(null);
    setShowModal(true);
  };

  const handleDelete = async (id) => {
    if (!window.confirm("Are you sure you want to delete this batch?")) return;

    try {
      await Tablemaintainanceapi.delete(`/batches/${id}`);
      fetchBatches();
      alert("Batch deleted successfully");
    } catch (error) {
      console.error("Error deleting batch:", error);
      alert("Error deleting batch");
    }
  };

  const handleSave = async () => {
    try {
      console.log("=== SAVING BATCH ===");

      if (modalMode === "add") {
        // Use batchService to send FormData with image
        await batchService.createBatch(currentBatch, batchLogoFile);
        alert("Batch created successfully");
      } else {
        await batchService.updateBatch(
          currentBatch.batchId,
          currentBatch,
          batchLogoFile,
        );
        alert("Batch updated successfully");
      }

      fetchBatches();
      setShowModal(false);
      setCurrentBatch(null);
      setBatchLogoFile(null);
    } catch (error) {
      console.error("Error saving batch:", error);
      alert(
        "Error saving batch: " +
          (error.response?.data?.message || error.message),
      );
    }
  };

  const filteredBatches = batches.filter((batch) =>
    batch.batchName?.toLowerCase().includes(searchTerm.toLowerCase()),
  );

  const totalPages = Math.ceil(filteredBatches.length / itemsPerPage);
  const indexOfLastItem = currentPage * itemsPerPage;
  const indexOfFirstItem = indexOfLastItem - itemsPerPage;
  const currentBatches = filteredBatches.slice(
    indexOfFirstItem,
    indexOfLastItem,
  );

  const handlePageChange = (pageNumber) => {
    setCurrentPage(pageNumber);
  };

  useEffect(() => {
    setCurrentPage(1);
  }, [searchTerm]);

  return (
    <div className="min-h-screen bg-gray-50 p-6">
      <div className="max-w-7xl mx-auto">
        {/* Header */}
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">
            Batch Management
          </h1>
          <p className="text-gray-600">
            Manage your course batches and schedules
          </p>
        </div>

        {/* Search and Add Section */}
        <div className="bg-white rounded-lg shadow-sm p-4 mb-6">
          <div className="flex flex-col sm:flex-row gap-4 items-center justify-between">
            <div className="relative flex-1 w-full">
              <Search className="absolute left-3 top-1/2 transform -translate-y-1/2 text-gray-400 w-5 h-5" />
              <input
                type="text"
                placeholder="Search batches..."
                value={searchTerm}
                onChange={(e) => setSearchTerm(e.target.value)}
                className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
              />
            </div>
            <button
              onClick={handleAdd}
              className="flex items-center gap-2 bg-blue-600 text-white px-4 py-2 rounded-lg hover:bg-blue-700 transition-colors whitespace-nowrap"
            >
              <Plus className="w-5 h-5" />
              Add Batch
            </button>
          </div>
        </div>

        {/* Table Section */}
        <div className="bg-white rounded-lg shadow-sm overflow-hidden">
          {loading ? (
            <div className="flex justify-center items-center py-12">
              <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-600"></div>
            </div>
          ) : currentBatches.length === 0 ? (
            <div className="text-center py-12 text-gray-500">
              <p className="text-lg">No batches found</p>
            </div>
          ) : (
            <>
              <div className="overflow-x-auto">
                <table className="w-full">
                  <thead className="bg-gray-50 border-b border-gray-200">
                    <tr>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Batch Name
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Course
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Duration
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Fees
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Logo
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Status
                      </th>
                      <th className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                        Actions
                      </th>
                    </tr>
                  </thead>
                  <tbody className="bg-white divide-y divide-gray-200">
                    {currentBatches.map((batch) => (
                      <tr key={batch.batchId} className="hover:bg-gray-50">
                        <td className="px-6 py-4 whitespace-nowrap">
                          <div className="text-sm font-medium text-gray-900">
                            {batch.batchName}
                          </div>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <div className="text-sm text-gray-900">
                            {batch.courseName}
                          </div>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <div className="text-sm text-gray-900">
                            {batch.batchStartDate && batch.batchEndDate
                              ? `${batch.batchStartDate} to ${batch.batchEndDate}`
                              : "N/A"}
                          </div>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <div className="text-sm text-gray-900">
                            {batch.courseFees ? `₹${batch.courseFees}` : "N/A"}
                          </div>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          {batch.batchLogoUrl && (
                            <img
                              src={`http://localhost:8080${batch.batchLogoUrl}`}
                              alt="batch logo"
                              className="w-12 h-12 object-cover rounded"
                            />
                          )}
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap">
                          <span
                            className={`px-2 inline-flex text-xs leading-5 font-semibold rounded-full ${
                              batch.batchIsActive
                                ? "bg-green-100 text-green-800"
                                : "bg-red-100 text-red-800"
                            }`}
                          >
                            {batch.batchIsActive ? "Active" : "Inactive"}
                          </span>
                        </td>
                        <td className="px-6 py-4 whitespace-nowrap text-sm font-medium">
                          <div className="flex gap-2">
                            <button
                              onClick={() => handleEdit(batch)}
                              className="text-blue-600 hover:text-blue-800 transition-colors"
                            >
                              <Edit2 className="w-5 h-5" />
                            </button>
                            <button
                              onClick={() => handleDelete(batch.batchId)}
                              className="text-red-600 hover:text-red-800 transition-colors"
                            >
                              <Trash2 className="w-5 h-5" />
                            </button>
                          </div>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>

              {filteredBatches.length > 0 && (
                <Pagination
                  currentPage={currentPage}
                  totalPages={totalPages}
                  onPageChange={handlePageChange}
                />
              )}
            </>
          )}
        </div>
      </div>

      {showModal && (
        <BatchModal
          mode={modalMode}
          batch={currentBatch}
          setBatch={setCurrentBatch}
          activeCourses={activeCourses}
          batchLogoFile={batchLogoFile}
          setBatchLogoFile={setBatchLogoFile}
          onSave={handleSave}
          onClose={() => {
            setShowModal(false);
            setCurrentBatch(null);
            setBatchLogoFile(null);
          }}
        />
      )}
    </div>
  );
};

// ============================================================
// Batch Modal Component
// ============================================================
const BatchModal = ({
  mode,
  batch,
  setBatch,
  activeCourses,
  batchLogoFile,
  setBatchLogoFile,
  onSave,
  onClose,
}) => {
  const handleChange = (e) => {
    const { name, value, type, checked } = e.target;
    setBatch((prev) => ({
      ...prev,
      [name]: type === "checkbox" ? checked : value,
    }));
  };

  return (
    <div className="fixed inset-0 bg-black bg-opacity-50 flex items-center justify-center p-4 z-50 overflow-y-auto">
      <div className="bg-white rounded-lg shadow-xl w-full max-w-4xl my-8">
        {/* Header */}
        <div className="flex justify-between items-center p-6 border-b border-gray-200">
          <h2 className="text-2xl font-bold text-gray-900">
            {mode === "add" ? "Add New Batch" : "Edit Batch"}
          </h2>
          <button
            onClick={onClose}
            className="text-gray-400 hover:text-gray-600 transition-colors"
          >
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Body */}
        <div className="p-6 max-h-[calc(100vh-200px)] overflow-y-auto">
          {/* Basic Information */}
          <div className="mb-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">
              Basic Information
            </h3>
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Batch Name *
                </label>
                <input
                  type="text"
                  name="batchName"
                  value={batch.batchName}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  required
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Course *
                </label>
                <select
                  name="courseId"
                  value={batch.courseId}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  required
                >
                  <option value="">Select Course</option>
                  {activeCourses.map((course) => (
                    <option key={course.courseId} value={course.courseId}>
                      {course.courseName}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Start Date
                </label>
                <input
                  type="date"
                  name="batchStartDate"
                  value={batch.batchStartDate}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  End Date
                </label>
                <input
                  type="date"
                  name="batchEndDate"
                  value={batch.batchEndDate}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Course Fees
                </label>
                <input
                  type="number"
                  name="courseFees"
                  value={batch.courseFees}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                  step="0.01"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Presentation Date
                </label>
                <input
                  type="datetime-local"
                  name="presentationDate"
                  value={batch.presentationDate}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <label className="block text-sm font-medium text-gray-700 mb-2">
                  Final Presentation Date
                </label>
                <input
                  type="datetime-local"
                  name="finalPresentationDate"
                  value={batch.finalPresentationDate}
                  onChange={handleChange}
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div className="flex items-center">
                <label className="flex items-center cursor-pointer">
                  <input
                    type="checkbox"
                    name="batchIsActive"
                    checked={batch.batchIsActive}
                    onChange={handleChange}
                    className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-blue-500"
                  />
                  <span className="ml-2 text-sm font-medium text-gray-700">
                    Active
                  </span>
                </label>
              </div>
            </div>
          </div>

          {/* Batch Logo Upload */}
          <div className="mb-6">
            <h3 className="text-lg font-semibold text-gray-900 mb-4">
              Batch Logo
            </h3>
            <div className="border-2 border-dashed border-gray-300 rounded-lg p-6 hover:border-blue-400 transition-colors">
              <input
                type="file"
                accept="image/*"
                onChange={(e) => setBatchLogoFile(e.target.files[0])}
                className="hidden"
                id="batch-logo-upload"
              />
              <label
                htmlFor="batch-logo-upload"
                className="cursor-pointer flex flex-col items-center gap-2"
              >
                <Upload className="w-12 h-12 text-gray-400" />
                <div className="text-sm text-gray-600">
                  <span className="text-blue-600 font-medium">
                    Click to upload
                  </span>{" "}
                  batch logo
                </div>
                <p className="text-xs text-gray-500">
                  PNG, JPG, GIF up to 10MB
                </p>
              </label>
            </div>

            {/* Preview */}
            {batchLogoFile && (
              <div className="mt-4 bg-gray-50 border border-gray-200 rounded-lg p-4">
                <div className="flex items-start gap-4">
                  <img
                    src={URL.createObjectURL(batchLogoFile)}
                    alt="preview"
                    className="w-24 h-24 object-cover rounded border border-gray-200"
                  />
                  <div className="flex-1">
                    <p className="text-sm font-medium text-gray-800">
                      New Logo Selected
                    </p>
                    <p className="text-xs text-gray-600 mt-1">
                      {batchLogoFile.name}
                    </p>
                    <p className="text-xs text-gray-500">
                      {(batchLogoFile.size / 1024).toFixed(2)} KB
                    </p>
                  </div>
                </div>
              </div>
            )}

            {/* Show current logo if editing */}
            {!batchLogoFile && batch.batchLogoUrl && (
              <div className="mt-4 bg-gray-50 border border-gray-200 rounded-lg p-4">
                <div className="flex items-start gap-4">
                  <img
                    src={`http://localhost:8080${batch.batchLogoUrl}`}
                    alt="current logo"
                    className="w-24 h-24 object-cover rounded border border-gray-200"
                  />
                  <div className="flex-1">
                    <p className="text-sm font-medium text-gray-800">
                      Current Logo
                    </p>
                    <p className="text-xs text-gray-500 break-all">
                      {batch.batchLogoUrl}
                    </p>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* Footer */}
        <div className="flex justify-end gap-3 p-6 border-t border-gray-200 bg-gray-50">
          <button
            onClick={onClose}
            className="px-4 py-2 text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
          >
            Cancel
          </button>
          <button
            onClick={onSave}
            className="flex items-center gap-2 px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
          >
            <Save className="w-5 h-5" />
            Save Batch
          </button>
        </div>
      </div>
    </div>
  );
};

export default BatchManagement;
