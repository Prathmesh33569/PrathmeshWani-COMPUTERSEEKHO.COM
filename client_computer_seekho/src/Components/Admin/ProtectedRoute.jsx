// import { Navigate } from "react-router-dom";

// const ProtectedRoute = ({ children }) => {
//   // const token = localStorage.getItem("token");

//   // if (!token) {
//   //   return <Navigate to="/admin/login" replace />;
//   // }

//   return children;
// };

// export default ProtectedRoute;

// import { Navigate } from "react-router-dom";
// import { jwtDecode } from "jwt-decode";

// const ProtectedRoute = ({ children }) => {
//   const token = localStorage.getItem("token");

//   if (!token) {
//     return <Navigate to="/admin/login" replace />;
//   }

//   try {
//     const decoded = jwtDecode(token);
//     const currentTime = Date.now() / 1000;
//     if (decoded.exp < currentTime) {
//       localStorage.removeItem("token");
//       return <Navigate to="/admin/login" replace />;
//     }

//     return children;
//   } catch (error) {
//     localStorage.removeItem("token");
//     return <Navigate to="/admin/login" replace />;
//   }
// };

// export default ProtectedRoute;

import { Navigate } from "react-router-dom";
import { useEffect, useState } from "react";

const ProtectedRoute = ({ children }) => {
  const [isAuth, setIsAuth] = useState(null);

  useEffect(() => {
    const token = localStorage.getItem("token");
    setIsAuth(!!token);
  }, []);

  if (isAuth === null) return null; // loading state

  if (!isAuth) {
    return <Navigate to="/admin/login" replace />;
  }

  return children;
};

export default ProtectedRoute;
