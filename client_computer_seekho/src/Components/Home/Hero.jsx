import React, { useEffect, useState } from "react";
import axios from "axios";
import heroImage from "../../assets/hero.png";

const Hero = () => {
  const [announcements, setAnnouncements] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    axios
      .get("http://localhost:8080/api/announcements/active")
      .then((response) => {
        setAnnouncements(response.data);
      })
      .catch((error) => {
        console.error("Error fetching announcements:", error);
      })
      .finally(() => {
        setLoading(false);
      });
  }, []);

  return (
    <div className="w-full flex flex-col items-center">
      {/* Announcement Bar */}
      <div className="w-full bg-sky-50 overflow-hidden">
        <div className="flex items-center gap-8 whitespace-nowrap px-4 py-2 animate-scroll-text min-w-max">
          {loading ? (
            <span className="text-sm font-semibold">
              Loading announcements...
            </span>
          ) : announcements.length > 0 ? (
            announcements.map((text, index) => (
              <span key={index} className="text-sm font-semibold text-gray-800">
                {text}
              </span>
            ))
          ) : (
            <span className="text-sm font-semibold">
              No announcements available
            </span>
          )}
        </div>
      </div>

      {/* Hero Image */}
      <div className="w-full flex justify-center py-6">
        <img src={heroImage} alt="hero" className="w-[60%] object-contain" />
      </div>
    </div>
  );
};

export default Hero;
