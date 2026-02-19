import { useEffect } from "react";
import { useNavigate } from "react-router-dom";

function OAuthSuccess() {
  const navigate = useNavigate();

  useEffect(() => {
    const params = new URLSearchParams(window.location.search);
    const token = params.get("token");
    const staff = params.get("name");

    if (token) {
      localStorage.setItem("token", token);
      localStorage.setItem("staff", JSON.stringify(staff));

      console.log("Token saved:", token);

      // ✅ FORCE APP RELOAD (IMPORTANT)
      window.location.href = "/admin";
    } else {
      window.location.href = "/admin/login";
    }
  }, []);

  return (
    <div className="min-h-screen flex items-center justify-center text-lg font-semibold">
      Logging in with Google...
    </div>
  );
}

export default OAuthSuccess;
