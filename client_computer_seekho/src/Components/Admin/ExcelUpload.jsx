import React, { useState } from "react";
import { Upload } from "lucide-react";
import axios from "axios";

const ExcelUpload = () => {
  const [selectedFile, setSelectedFile] = useState(null);
  const [loading, setLoading] = useState(false);

  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    if (!file.name.endsWith(".xls") && !file.name.endsWith(".xlsx")) {
      alert("Please select a valid Excel file (.xls or .xlsx)");
      e.target.value = "";
      return;
    }

    setSelectedFile(file);
  };

  const handleSubmit = async () => {
    if (!selectedFile) return;

    const formData = new FormData();
    formData.append("file", selectedFile); // 👈 must match backend @RequestParam("file")

    try {
      setLoading(true);

      const response = await axios.post(
        "http://localhost:8080/api/albums/upload-excel",
        formData,
        {
          headers: {
            "Content-Type": "multipart/form-data",
          },
        },
      );

      alert(response.data); // success message from backend
      setSelectedFile(null);
    } catch (error) {
      console.error("Upload error:", error);
      alert("Failed to upload Excel file");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="bg-white rounded-lg shadow-sm p-6">
      <div className="mb-6">
        <h2 className="text-2xl font-bold text-gray-900">Excel Upload</h2>
        <p className="text-gray-500">Upload master data using Excel files</p>
      </div>

      <div className="max-w-md space-y-4">
        <div>
          <label className="block text-sm font-medium text-gray-700 mb-1">
            Select Excel File
          </label>
          <input
            type="file"
            accept=".xls,.xlsx"
            onChange={handleFileChange}
            className="w-full px-4 py-2 border border-gray-300 rounded-lg focus:ring-2 focus:ring-blue-500"
          />
        </div>

        {selectedFile && (
          <p className="text-sm text-gray-600">
            Selected File:{" "}
            <span className="font-medium">{selectedFile.name}</span>
          </p>
        )}

        <button
          onClick={handleSubmit}
          disabled={!selectedFile || loading}
          className={`w-full md:w-auto flex items-center justify-center gap-2 px-6 py-2 rounded-lg transition-colors
            ${
              selectedFile
                ? "bg-blue-600 text-white hover:bg-blue-700"
                : "bg-gray-400 text-white cursor-not-allowed"
            }`}
        >
          <Upload className="w-5 h-5" />
          {loading ? "Uploading..." : "Submit"}
        </button>
      </div>
    </div>
  );
};

export default ExcelUpload;
