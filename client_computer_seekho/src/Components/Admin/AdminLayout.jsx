import { Outlet, NavLink, useNavigate } from "react-router-dom";
import { useState } from "react";
import { ClipboardList, Users, Wrench, FileSpreadsheet } from "lucide-react"; // ✅ Import icons
import AdminHeader from "./AdminHeader";
import AdminFooter from "./AdminFooter";

function AdminLayout() {
  const navigate = useNavigate();
  const [isSidebarOpen, setIsSidebarOpen] = useState(true);

  const handleLogout = () => {
    localStorage.removeItem("isAuthenticated");
    navigate("/admin/login");
  };
  
  const menuItems = [
    { path: "/admin/follow-up", label: "Follow Up", icon: ClipboardList },
    { path: "/admin/registered-students", label: "Registered Students", icon: Users }, // ✅ NEW
    { path: "/admin/table-maintenance", label: "Table Maintenance", icon: Wrench },
    { path: "/admin/excel-upload", label: "Excel Upload", icon: FileSpreadsheet },
  ];

  return (
    <div className="flex h-screen bg-gray-100">
      {/* Sidebar */}
      <aside
        className={`${isSidebarOpen ? "w-64" : "w-20"} bg-gray-800 text-white transition-all duration-300`}
      >
        <div className="p-4 flex justify-between items-center border-b border-gray-700">
          {isSidebarOpen && <h1 className="text-xl font-bold">Admin Panel</h1>}
          <button
            onClick={() => setIsSidebarOpen(!isSidebarOpen)}
            className="text-2xl hover:bg-gray-700 p-2 rounded"
          >
            {isSidebarOpen ? "◀" : "▶"}
          </button>
        </div>

        <nav className="mt-6">
          {menuItems.map((item) => {
            const Icon = item.icon; // ✅ Get icon component
            return (
              <NavLink
                key={item.path}
                to={item.path}
                className={({ isActive }) =>
                  `flex items-center px-4 py-3 hover:bg-gray-700 transition ${
                    isActive ? "bg-gray-700 border-l-4 border-blue-500" : ""
                  }`
                }
              >
                <Icon size={24} /> {/* ✅ Render icon */}
                {isSidebarOpen && <span className="ml-4">{item.label}</span>}
              </NavLink>
            );
          })}
        </nav>
      </aside>

      {/* ✅ MAIN CONTENT AREA */}
      <div className="flex flex-col flex-1 bg-gray-50">
        <AdminHeader />

        <div className="flex-1 overflow-y-auto p-6">
          <div className="bg-white rounded-xl shadow-sm p-6">
            <Outlet />
          </div>
        </div>

        <AdminFooter />
      </div>
    </div>
  );
}

export default AdminLayout;