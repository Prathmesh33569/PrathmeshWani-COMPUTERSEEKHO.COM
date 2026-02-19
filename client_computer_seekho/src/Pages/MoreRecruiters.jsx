import { useEffect, useState } from "react";
import axios from "axios";

const MoreRecruiters = () => {
  const [logos, setLogos] = useState([]);
  const [loading, setLoading] = useState(true);
// const navigate = useNavigate();

  useEffect(() => {
  axios
    .get("http://localhost:8080/api/recruiters/images")
    .then((response) => {
      setLogos(response.data);   // 👈 Axios auto JSON parse karta hai
      setLoading(false);
    })
    .catch((error) => {
      console.error("Error fetching recruiter images:", error);
      setLoading(false);
    });
}, []);

  if (loading) {
    return <p className="text-center">Loading recruiters...</p>;
  }

   return (
    <section className="py-12 text-center">
      <h2 className="text-2xl font-semibold text-blue-900 mb-8">
        Major Recruiters
      </h2>

      <div className="max-w-5xl mx-auto grid grid-cols-2 sm:grid-cols-3 md:grid-cols-4 gap-8 place-items-center">
        {logos.map((path, index) => (
          <img
            key={index}
            src={path}
            alt="Recruiter"
            className="h-12 grayscale hover:grayscale-0 transition"
            onError={(e) => {
              e.target.src = "/assets/Recruiters/default.png";
            }}
          />
        ))}
      </div>
    </section>
  );


};

export default MoreRecruiters;