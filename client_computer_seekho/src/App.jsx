import { Routes, Route, Navigate } from "react-router-dom";
import Header from "./Components/Header/Header";
import Home from "./Pages/Home";
import Placement from "./Pages/Placement";
import Courses from "./Pages/Courses";
import CampusLife from "./Pages/CampusLife";
import Faculty from "./Pages/Faculty";
import OPDC from "./Pages/OPDC";
import ContactSection from "./Pages/ContactSection";
import PlacementAlbum from "./Pages/PlacementAlbum";
import MoreRecruiters from "./Pages/MoreRecruiters";
import CoursesDetails from "./Pages/CoursesDetails";
import PublicLayout from "./Components/Layout/PublicLayout";
import Login from "./Components/Admin/Login";
import ProtectedRoute from "./Components/Admin/ProtectedRoute";
import AdminLayout from "./Components/Admin/AdminLayout";
import TableMaintenance from "./Components/Admin/TableMaintanance/TableMaintenance";
import ExcelUpload from "./Components/Admin/ExcelUpload";
import Signup from "./Components/Admin/Signup";
import OAuthSuccess from "./Components/Admin/OAuthSuccess";
import FollowUpDashboard from "./Pages/FollowUpDashboard";
import RegisteredStudentsPage from './Components/Admin/RegisteredStudents/RegisteredStudentsPage';

function App() {
  return (
    <>
      <Routes>
        <Route element={<PublicLayout />}>
          <Route path="/" element={<Home />} />
          <Route path="/placement" element={<Placement />} />
          <Route path="/placement/:albumId" element={<PlacementAlbum />} />
          <Route path="/courses" element={<Courses />} />
          <Route path="/campus-life" element={<CampusLife />} />
          {/* <Route path="/pg-diploma" element={<OPDC />} /> */}
          <Route path="/faculty" element={<Faculty />} />
          <Route path="/contact" element={<ContactSection />} />
          <Route path="/more-recruiters" element={<MoreRecruiters />} />
          <Route path="/courses/:slug" element={<CoursesDetails />} />
        </Route>

        <Route path="/admin/login" element={<Login />} />
        <Route path="/admin/signup" element={<Signup />} />
        <Route path="/oauth-success" element={<OAuthSuccess />} />

        <Route
          path="/admin"
          element={
            <ProtectedRoute>
              <AdminLayout />
            </ProtectedRoute>
          }
        >
          <Route index element={<Navigate to="follow-up" replace />} />
          <Route path="follow-up" element={<FollowUpDashboard/>} />
          <Route path="table-maintenance" element={<TableMaintenance />} />
          <Route path="excel-upload" element={<ExcelUpload />} />
          // Add this route:
          <Route path="/admin/registered-students" element={<RegisteredStudentsPage />} />
        </Route>
      </Routes>
    </>
  );
}

export default App;
