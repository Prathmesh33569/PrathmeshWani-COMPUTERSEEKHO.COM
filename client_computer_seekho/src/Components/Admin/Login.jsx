import React, { useState } from "react";
import { useNavigate, Link } from "react-router-dom";
import authService from "../../Services/authService";

function Login() {
  const navigate = useNavigate();
  const [formData, setFormData] = useState({
    staffUsername: "",
    staffPassword: "",
  });
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  const handleGoogleLogin = () => {
    window.location.href = "http://localhost:8080/oauth2/authorization/google";
  };

  const handleSubmit = async (e) => {
    e.preventDefault();
    setError("");
    setLoading(true);

    try {
      const response = await authService.login(
        formData.staffUsername,
        formData.staffPassword,
      );

      console.log("Login successful:", response);
      navigate("/admin");
    } catch (err) {
      console.error("Login error:", err);
      setError(
        err.response?.data?.error ||
          err.response?.data?.message ||
          "Invalid username or password",
      );
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen flex items-center justify-center bg-gradient-to-br from-blue-600 via-blue-700 to-blue-800 px-4">
      <div className="bg-white rounded-2xl shadow-2xl p-8 w-full max-w-md animate-fadeIn">
        <div className="text-center mb-8">
          <h2 className="text-3xl font-bold text-gray-800 mb-2">Staff Login</h2>
          <p className="text-gray-600 text-sm">
            Computer Seekho - Staff Portal
          </p>
        </div>

        {error && (
          <div className="bg-red-50 border-l-4 border-red-500 text-red-700 p-4 mb-6 rounded">
            <p className="text-sm">{error}</p>
          </div>
        )}

        <form onSubmit={handleSubmit} className="space-y-6">
          <div>
            <label
              htmlFor="staffUsername"
              className="block text-sm font-semibold text-gray-700 mb-2"
            >
              Username
            </label>
            <input
              type="text"
              id="staffUsername"
              name="staffUsername"
              value={formData.staffUsername}
              onChange={handleChange}
              placeholder="Enter your username"
              required
              disabled={loading}
              className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100 disabled:cursor-not-allowed"
            />
          </div>

          <div>
            <label
              htmlFor="staffPassword"
              className="block text-sm font-semibold text-gray-700 mb-2"
            >
              Password
            </label>
            <input
              type="password"
              id="staffPassword"
              name="staffPassword"
              value={formData.staffPassword}
              onChange={handleChange}
              placeholder="Enter your password"
              required
              disabled={loading}
              className="w-full px-4 py-3 border-2 border-gray-300 rounded-lg focus:ring-2 focus:ring-purple-500 focus:border-transparent transition disabled:bg-gray-100 disabled:cursor-not-allowed"
            />
          </div>

          <button
            type="submit"
            disabled={loading}
            className="w-full bg-gradient-to-r from-blue-900 to-indigo-700 text-white font-semibold py-3 rounded-lg hover:from-purple-700 hover:to-indigo-700 transition transform hover:scale-[1.02] disabled:opacity-60 disabled:cursor-not-allowed disabled:transform-none shadow-lg"
          >
            {loading ? "Logging in..." : "Login"}
          </button>
          <button
            onClick={handleGoogleLogin}
            className="w-full mt-4 bg-white border border-gray-300 text-gray-700 font-semibold py-3 rounded-lg hover:bg-gray-100 transition flex items-center justify-center gap-2 shadow"
          >
            <img
              src="https://developers.google.com/identity/images/g-logo.png"
              alt="Google"
              className="w-5 h-5"
            />
            Login with Google
          </button>
        </form>

        <p className="text-center text-gray-600 text-sm mt-6">
          Don't have an account?{" "}
          <Link
            to="/admin/signup"
            className="text-purple-600 font-semibold hover:text-purple-700 transition"
          >
            Sign up here
          </Link>
        </p>
      </div>
    </div>
  );
}

export default Login;
