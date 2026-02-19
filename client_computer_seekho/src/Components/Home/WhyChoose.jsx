import React from "react";
import { FaCheckCircle } from "react-icons/fa";
import whyLogo from "../../assets/Whylogo.png";

function WhyChoose() {
  return (
    <section className="bg-[#3f5d7d] text-white py-20">
      <div className="max-w-7xl mx-auto flex items-center gap-16 px-10">
        <div className="w-1/2 flex justify-center">
          <img src={whyLogo} alt="" className="w-[80%]" />
        </div>

        <div className="w-1/2">
          <h2 className="text-4xl font-bold mb-6">
            Why Choose <span className="text-red-500">Computer Seekho</span> ?
          </h2>

          <p className="text-lg leading-relaxed mb-8">
            Our institute has been present for over 20 years in the market. We
            make the most of all our students.
          </p>

          <ul className="space-y-5">
            <li className="flex items-center gap-4 text-lg font-semibold">
              <FaCheckCircle className="text-red-500 text-2xl" />
              Best in class Infrastructure
            </li>

            <li className="flex items-center gap-4 text-lg font-semibold">
              <FaCheckCircle className="text-red-500 text-2xl" />
              Best Faculty / Teachers
            </li>

            <li className="flex items-center gap-4 text-lg font-semibold">
              <FaCheckCircle className="text-red-500 text-2xl" />
              Best Learning Methodology
            </li>

            <li className="flex items-center gap-4 text-lg font-semibold">
              <FaCheckCircle className="text-red-500 text-2xl" />
              More than 95% Placement for 10 Consecutive batches
            </li>
          </ul>
        </div>
      </div>
    </section>
  );
}

export default WhyChoose;
