import { Link } from "react-router-dom";
import Testimonials from "./Testimonials";

const OPDC = () => {
  return (
    <section className="max-w-[1300px] mx-auto px-6 lg:px-20 py-16">
        <div className="grid grid-cols-1 lg:grid-cols-2 gap-16 items-start">
        
        {/* LEFT CONTENT */}
        <div className="lg:pl-10 xl:pl-16">

          
          <h1 className="text-4xl font-bold text-[#0b3c6d] mb-6">
            Turning Adversity into an Advantage
          </h1>

          <p className="text-gray-600 mb-6 leading-relaxed">
            We may have been compelled to move our training on-line but we were
            unwilling to sacrifice quality.
          </p>

          <p className="text-gray-600 mb-6 leading-relaxed">
            Our team was determined to provide high-quality, engaging learning
            experiences for our students with focus on creating three form of
            interaction for students in the online environment.
          </p>

          <p className="text-gray-700 font-medium mb-8">
            The success of this approach is evident from our students
          </p>

          {/* LINKS */}
          <ul className="space-y-4 text-blue-600 font-medium">
            <li>
              <Link to="#" className="hover:underline">
                Testimonials : PG-DAC & PG-DBDA Feb 2020
              </Link>
            </li>
            <li>
              <Link to="#" className="hover:underline">
                Testimonials : e-DAC & e-DBDA Sep 2020
              </Link>
            </li>
            <li>
              <Link to="#" className="hover:underline">
                Online Interactions with Industry Experts
              </Link>
            </li>
            <li>
              <Link to="#" className="hover:underline">
                Guidelines & Support from Seniors
              </Link>
            </li>
            <li>
              <Link to="#" className="hover:underline">
                Online Competition
              </Link>
            </li>
            <li>
              <Link to="#" className="hover:underline">
                Check Online Celebration Video
              </Link>
            </li>
          </ul>
        </div>

       {/* RIGHT IMAGE */}
<div className="flex justify-center lg:justify-end overflow-hidden">
  <img
    src="/image.png"
    alt="Student Interaction Model"
    className="w-full max-w-[420px] lg:max-w-[520px] h-auto object-contain"
  />
</div>


      </div>
      <Testimonials />

    </section>

  );
};

export default OPDC;
