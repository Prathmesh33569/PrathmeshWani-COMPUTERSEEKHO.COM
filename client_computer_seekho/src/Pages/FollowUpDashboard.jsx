import React, { useState, useEffect } from "react";
import Header from "../Components/Admin/common/Header";
import Navigation from "../Components/Admin/common/Navigation";
import SearchFilter from "../Components/Admin/common/SearchFilter";
import Loading from "../Components/Admin/common/Loading";
import EnquiryStats from "../Components/Admin/enquiry/EnquiryStats";
import EnquiryList from "../Components/Admin/enquiry/EnquiryList";
import CallModal from "../Components/Admin/enquiry/CallModal";
import AddEnquiryModal from "../Components/Admin/enquiry/AddEnquiryModal";
import EditEnquiryModal from "../Components/Admin/enquiry/EditEnquiryModal";
import RegisterStudentModal from "../Components/Admin/enquiry/RegisterStudentModal";
import {
  getFollowupsForStaff,
  getAllFollowups,
} from "../Services/enquiryService";
import {
  LOGGED_IN_STAFF_ID,
  VIEW_TYPES,
  FILTER_OPTIONS,
} from "../utils/constants";
import { isPending, isToday } from "../utils/helpers";

const FollowUpDashboard = () => {
  const [activeView, setActiveView] = useState(VIEW_TYPES.MY_FOLLOWUPS);
  const [followups, setFollowups] = useState([]);
  const [loading, setLoading] = useState(false);
  const [selectedEnquiry, setSelectedEnquiry] = useState(null);
  const [showCallModal, setShowCallModal] = useState(false);
  const [showAddEnquiry, setShowAddEnquiry] = useState(false);
  const [showEditEnquiry, setShowEditEnquiry] = useState(false);
  const [showRegisterStudent, setShowRegisterStudent] = useState(false);
  const [editEnquiryId, setEditEnquiryId] = useState(null);
  const [registerEnquiryId, setRegisterEnquiryId] = useState(null);
  const [searchTerm, setSearchTerm] = useState("");
  const [filterStatus, setFilterStatus] = useState(FILTER_OPTIONS.ALL);

  useEffect(() => {
    fetchFollowups();
  }, [activeView]);

  const fetchFollowups = async () => {
    setLoading(true);
    try {
      let data;
      if (activeView === VIEW_TYPES.MY_FOLLOWUPS) {
        console.log("Sachin", LOGGED_IN_STAFF_ID);
        data = await getFollowupsForStaff(LOGGED_IN_STAFF_ID);
      } else {
        data = await getAllFollowups();
      }
      setFollowups(data);
    } catch (error) {
      console.error("Error fetching followups:", error);
      alert("Error fetching follow-ups. Please check if backend is running.");
    } finally {
      setLoading(false);
    }
  };

  const handleCallClick = (enquiry) => {
    setSelectedEnquiry(enquiry);
    setShowCallModal(true);
  };

  const handleEditClick = (enquiry) => {
    setEditEnquiryId(enquiry.enquiryId);
    setShowEditEnquiry(true);
  };

  const handleRegisterClick = (enquiry) => {
    setRegisterEnquiryId(enquiry.enquiryId);
    setShowRegisterStudent(true);
  };

  const handleCloseCallModal = () => {
    setShowCallModal(false);
    setSelectedEnquiry(null);
  };

  const handleCloseAddEnquiry = () => {
    setShowAddEnquiry(false);
  };

  const handleCloseEditEnquiry = () => {
    setShowEditEnquiry(false);
    setEditEnquiryId(null);
  };

  const handleCloseRegisterStudent = () => {
    setShowRegisterStudent(false);
    setRegisterEnquiryId(null);
  };

  const handleSuccess = () => {
    fetchFollowups();
  };

  // Filter followups
  const filteredFollowups = followups.filter((f) => {
    const matchesSearch =
      f.enquirerName?.toLowerCase().includes(searchTerm.toLowerCase()) ||
      f.enquirerMobile?.toString().includes(searchTerm);

    if (filterStatus === FILTER_OPTIONS.PENDING) {
      return matchesSearch && isPending(f.followupDate);
    }
    if (filterStatus === FILTER_OPTIONS.TODAY) {
      return matchesSearch && isToday(f.followupDate);
    }
    return matchesSearch;
  });

  return (
    <div className="min-h-screen bg-gray-50">
      <Header />

      <Navigation
        activeView={activeView}
        onViewChange={setActiveView}
        onAddEnquiry={() => setShowAddEnquiry(true)}
      />

      <div className="max-w-7xl mx-auto px-4 py-4">
        <EnquiryStats followups={followups} />
      </div>

      <div className="max-w-7xl mx-auto px-4 py-2">
        <div className="mb-6">
          <SearchFilter
            searchTerm={searchTerm}
            onSearchChange={setSearchTerm}
            filterStatus={filterStatus}
            onFilterChange={setFilterStatus}
          />
        </div>

        {loading ? (
          <Loading message="Loading follow-ups..." />
        ) : (
          <EnquiryList
            enquiries={filteredFollowups}
            onCallClick={handleCallClick}
            onEditClick={handleEditClick}
            onRegisterClick={handleRegisterClick}
            loading={loading}
          />
        )}
      </div>

      {showCallModal && selectedEnquiry && (
        <CallModal
          enquiry={selectedEnquiry}
          onClose={handleCloseCallModal}
          onSuccess={handleSuccess}
        />
      )}

      {showAddEnquiry && (
        <AddEnquiryModal
          onClose={handleCloseAddEnquiry}
          onSuccess={handleSuccess}
        />
      )}

      {showEditEnquiry && editEnquiryId && (
        <EditEnquiryModal
          enquiryId={editEnquiryId}
          onClose={handleCloseEditEnquiry}
          onSuccess={handleSuccess}
        />
      )}

      {showRegisterStudent && registerEnquiryId && (
        <RegisterStudentModal
          enquiryId={registerEnquiryId}
          onClose={handleCloseRegisterStudent}
          onSuccess={handleSuccess}
        />
      )}
    </div>
  );
};

export default FollowUpDashboard;
