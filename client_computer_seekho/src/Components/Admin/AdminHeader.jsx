// // import appLogo3 from "../../assets/CSlogo3.jpg";

// // function AdminHeader() {
// //   return (
// //     <header className="bg-white shadow-sm border-b px-6 py-2 flex items-center justify-between">
// //       {/* Left: Logo */}
// //       <div className="flex items-center gap-3">
// //         <img
// //           src={appLogo3}
// //           alt="Computer Seekho Logo"
// //           className="h-12 object-contain"
// //         />
// //         <span className="text-lg font-semibold text-gray-700">
// //           Computer Seekho
// //         </span>
// //       </div>

// //       {/* Right: Admin Info */}
// //       <div className="flex items-center gap-4">
// //         <span className="text-sm text-gray-600">Welcome, Admin</span>
// //         <div className="w-9 h-9 bg-green-600 rounded-full flex items-center justify-center text-white font-bold">
// //           A
// //         </div>
// //       </div>
// //     </header>
// //   );
// // }

// // export default AdminHeader;
// import appLogo3 from "../../assets/CSlogo3.jpg";

// function AdminHeader() {
//   const staffData = JSON.parse(localStorage.getItem("staff"));

//   const staffName = staffData?.staffName || "Admin";

//   const firstLetter = staffName.charAt(0).toUpperCase();

//   return (
//     <header className="bg-white shadow-sm border-b px-6 py-2 flex items-center justify-between relative">
//       <div className="flex items-center gap-3">
//         <img
//           src={appLogo3}
//           alt="Computer Seekho Logo"
//           className="h-12 object-contain"
//         />
//       </div>

//       <div className="absolute left-1/2 transform -translate-x-1/2">
//         <span className="text-xl font-bold text-gray-700">Computer Seekho</span>
//       </div>

//       <div className="flex items-center gap-4">
//         <span className="text-sm text-gray-600">Welcome, {staffName}</span>

//         <div className="w-9 h-9 bg-red-600 rounded-full flex items-center justify-center text-white font-bold">
//           {firstLetter}
//         </div>
//       </div>
//     </header>
//   );
// }

// export default AdminHeader;

import { useState } from "react";
import { useNavigate } from "react-router-dom";
import appLogo3 from "../../assets/CSlogo3.jpg";

function AdminHeader() {
  const navigate = useNavigate();
  const [showMenu, setShowMenu] = useState(false);

  const staffData = JSON.parse(localStorage.getItem("staff"));
  const staffName = staffData?.staffName || "Admin";
  const firstLetter = staffName.charAt(0).toUpperCase();

  const handleLogout = () => {
    localStorage.removeItem("staff"); // clear staff data
    localStorage.removeItem("token"); // if you store JWT token
    navigate("/admin/login"); // redirect to login page
  };

  return (
    <header className="bg-white shadow-sm border-b px-6 py-2 flex items-center justify-between relative">
      {/* Left Logo */}
      <div className="flex items-center gap-3">
        <img
          src={appLogo3}
          alt="Computer Seekho Logo"
          className="h-12 object-contain"
        />
      </div>

      {/* Center Title */}
      <div className="absolute left-1/2 transform -translate-x-1/2">
        <span className="text-xl font-bold text-gray-700">Computer Seekho</span>
      </div>

      {/* Right Admin Section */}
      <div className="relative flex items-center gap-3">
        <span className="text-sm text-gray-600">Welcome, {staffName}</span>

        {/* Avatar */}
        <div
          onClick={() => setShowMenu(!showMenu)}
          className="w-9 h-9 bg-red-600 rounded-full flex items-center justify-center text-white font-bold cursor-pointer hover:bg-red-700 transition"
        >
          {firstLetter}
        </div>

        {/* Dropdown Menu */}
        {showMenu && (
          <div className="absolute right-0 top-12 bg-white shadow-lg rounded-lg border w-32">
            <button
              onClick={handleLogout}
              className="w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-red-50 rounded-lg transition"
            >
              Logout
            </button>
          </div>
        )}
      </div>
    </header>
  );
}

export default AdminHeader;
