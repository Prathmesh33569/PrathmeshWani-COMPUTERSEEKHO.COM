import React from "react";
import {
  User,
  Phone,
  BookOpen,
  Calendar,
  CreditCard,
  FileText,
  Download,
  Send,
} from "lucide-react";
import { downloadReceipt, sendReceipt } from "../../../Services/receiptService";
import { useState } from "react";

const RegisteredStudentCard = ({ student, onRefresh }) => {
  // ✅ ADD ERROR HANDLING - Check if student exists
  const [sending, setSending] = useState(false);

  if (!student) {
    console.error("Student object is undefined");
    return (
      <div className="bg-red-50 border border-red-200 rounded-lg p-4">
        <p className="text-red-600">Error: Student data not available</p>
      </div>
    );
  }

  // ✅ LOG STUDENT DATA FOR DEBUGGING
  console.log("Student data:", student);

  const handleDownloadReceipt = async () => {
    if (!student.receiptId) {
      alert("No receipt available for this student");
      return;
    }

    try {
      // TODO: Your friend will implement this
      await downloadReceipt(student.receiptId);
      alert("Receipt download feature will be implemented by your friend");
    } catch (error) {
      console.error("Error downloading receipt:", error);
      alert("Download feature coming soon!");
    }
  };

  // const handleSendReceipt = async () => {
  //   if (!student.receiptId) {
  //     alert('No receipt available for this student');
  //     return;
  //   }

  //   // Prompt for email (your friend can modify this)
  //   const email = prompt('Enter email address to send receipt:');
  //   if (!email) return;

  //   try {
  //     // TODO: Your friend will implement this
  //     await sendReceipt(student.receiptId, email);
  //     alert('Receipt email feature will be implemented by your friend');
  //   } catch (error) {
  //     console.error('Error sending receipt:', error);
  //     alert('Send feature coming soon!');
  //   }
  // };

  const handleSendReceipt = async () => {
    if (!student.receiptId) {
      alert("No receipt available for this student");
      return;
    }

    try {
      setSending(true);
      const res = await sendReceipt(student.receiptId);
      alert(res); // backend message
    } catch (error) {
      console.error("Error sending receipt:", error);
      alert("Failed to send receipt email!");
    } finally {
      setSending(false);
    }
  };

  const formatDate = (date) => {
    if (!date) return "N/A";
    return new Date(date).toLocaleDateString("en-IN");
  };

  return (
    <div className="bg-white rounded-lg shadow hover:shadow-md transition p-4 border-l-4 border-purple-500">
      <div className="flex items-start justify-between">
        <div className="flex-1 grid grid-cols-5 gap-4">
          {/* Column 1: Student Info */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <User size={16} className="text-gray-500" />
              <span className="font-semibold text-gray-800">
                {student.studentName || "N/A"}
              </span>
            </div>
            <div className="flex items-center gap-2 text-sm text-gray-600 mb-1">
              <Phone size={14} />
              <span>{student.studentMobile || "N/A"}</span>
            </div>
            <div className="text-xs text-gray-500">
              ID: #{student.studentId || "N/A"}
            </div>
          </div>

          {/* Column 2: Course & Batch */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <BookOpen size={16} className="text-blue-600" />
              <span className="font-medium text-gray-700">
                Course ID: {student.courseId || "N/A"}
              </span>
            </div>
            <div className="text-sm text-gray-600">
              Batch ID: {student.batchId || "N/A"}
            </div>
          </div>

          {/* Column 3: Payment Info */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <CreditCard size={16} className="text-green-600" />
              <span className="font-semibold text-green-700">
                ₹{student.amount || "0"}
              </span>
            </div>
            <div className="text-sm text-gray-600">
              {student.paymentTypeDesc || "N/A"}
            </div>
            <div className="text-xs text-gray-500">
              {formatDate(student.paymentDate)}
            </div>
          </div>

          {/* Column 4: Receipt Info */}
          <div>
            <div className="flex items-center gap-2 mb-2">
              <FileText size={16} className="text-purple-600" />
              <span className="font-medium text-gray-700">
                Receipt #{student.receiptId || "N/A"}
              </span>
            </div>
            <div className="text-sm text-gray-600">
              Date: {formatDate(student.receiptDate)}
            </div>
            <div className="text-xs text-gray-500">
              Registered: {formatDate(student.registeredAt)}
            </div>
          </div>

          {/* Column 5: Action Buttons */}
          <div className="flex items-center justify-end gap-2">
            <button
              onClick={handleDownloadReceipt}
              disabled={!student.receiptId}
              className="px-4 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition flex items-center gap-2 font-medium disabled:bg-gray-400 disabled:cursor-not-allowed"
              title="Download Receipt"
            >
              <Download size={18} />
              Download
            </button>
            <button
              onClick={handleSendReceipt}
              disabled={!student.receiptId || sending}
              className={`px-4 py-2 rounded-lg transition flex items-center gap-2 font-medium 
    ${
      sending
        ? "bg-gray-400 cursor-not-allowed"
        : "bg-green-600 text-white hover:bg-green-700"
    }`}
              title="Send Receipt via Email"
            >
              {sending ? (
                <>
                  <span className="animate-spin">⏳</span>
                  Sending...
                </>
              ) : (
                <>
                  <Send size={18} />
                  Send
                </>
              )}
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default RegisteredStudentCard;
