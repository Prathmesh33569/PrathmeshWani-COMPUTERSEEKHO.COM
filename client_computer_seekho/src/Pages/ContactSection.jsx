import React, { useState } from "react";

const ContactSection = () => {
  const [formData, setFormData] = useState({
    name: "",
    email: "",
    message: "",
  });

  const [loading, setLoading] = useState(false); // ✅ loading state

  const handleChange = (e) => {
    setFormData({
      ...formData,
      [e.target.name]: e.target.value,
    });
  };

  const handleSubmit = async (e) => {
    e.preventDefault();

    setLoading(true); // ✅ start loading

    try {
      const response = await fetch("http://localhost:8080/api/email", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
        },
        body: JSON.stringify(formData),
      });

      const result = await response.text();
      alert(result);

      // clear form after success
      setFormData({
        name: "",
        email: "",
        message: "",
      });
    } catch (error) {
      console.error("Error:", error);
      alert("Failed to send message");
    } finally {
      setLoading(false); // ✅ stop loading
    }
  };

  return (
    <div className="max-w-7xl mx-auto px-6 py-12">
      <div className="grid md:grid-cols-2 gap-10">
        <div>
          <h2 className="text-2xl font-bold text-blue-900 mb-4">Our Origin</h2>
          <p className="text-gray-700 leading-relaxed">
            Computer Seekho is a modern IT training institute dedicated to
            empowering students with practical technical skills. Our mission is
            to bridge the gap between academic learning and industry
            requirements by providing hands-on training in the latest
            technologies. We focus on building strong foundations, real-world
            project experience, and career-oriented learning to help students
            achieve their professional goals in the IT industry.
          </p>
        </div>

        <div>
          <h2 className="text-2xl font-bold text-blue-900 mb-4">Reach us at</h2>
          <p className="text-gray-700">
            <strong>Computer Seekho - IT Training Institute</strong>
            <br />
            2nd Floor, Computer Seekho Learning Center,
            <br />
            Andheri (West), Mumbai, Maharashtra - 400058
            <br />
            India
            <br />
            <br />
            <strong>Mobile:</strong> +91 98765 43210 / +91 9503369902
            <br />
            <strong>Email:</strong> support@sheekhocomputer@gmail.com
          </p>
        </div>
      </div>

      <div className="grid md:grid-cols-2 gap-10 mt-12">
        <div className="w-full h-[350px] border rounded-lg overflow-hidden">
          <iframe
            title="SM VITA Location"
            className="w-full h-full"
            src="https://www.google.com/maps/embed?pb=!1m18!1m12!1m3!1d3769.906025253093!2d72.83000797442975!3d19.11177815084987!2m3!1f0!2f0!3f0!3m2!1i1024!2i768!4f13.1!3m3!1m2!1s0x3be7c9c3a5e26d7b%3A0x89a89f343cff9c29!2sSM%20VITA!5e0!3m2!1sen!2sin!4v1768888547731"
            loading="lazy"
            referrerPolicy="no-referrer-when-downgrade"
          ></iframe>
        </div>

        <div>
          <h2 className="text-2xl font-bold text-blue-900 mb-4">
            Connect With Computer Seekho
          </h2>

          <form className="space-y-4" onSubmit={handleSubmit}>
            <input
              type="text"
              name="name"
              placeholder="Name*"
              className="w-full border px-4 py-2"
              value={formData.name}
              onChange={handleChange}
              required
            />

            <input
              type="email"
              name="email"
              placeholder="Email"
              className="w-full border px-4 py-2"
              value={formData.email}
              onChange={handleChange}
              required
            />

            <textarea
              name="message"
              placeholder="Message"
              rows="5"
              className="w-full border px-4 py-2"
              value={formData.message}
              onChange={handleChange}
              required
            ></textarea>

            <button
              type="submit"
              disabled={loading}
              className={`px-6 py-2 text-white ${
                loading
                  ? "bg-gray-400 cursor-not-allowed"
                  : "bg-red-600 hover:bg-red-700"
              }`}
            >
              {loading ? "Sending..." : "SEND MESSAGE"}
            </button>
          </form>
        </div>
      </div>
    </div>
  );
};

export default ContactSection;
