import {
  MapPin,
  Phone,
  Mail,
  Facebook,
  Linkedin,
  Youtube,
  ChevronUp,
} from "lucide-react";
import { Link } from "react-router-dom";

const Footer = () => {
  const links = [
    { name: "Home", path: "/" },
    { name: "Placement", path: "/placement" },
    { name: "Campus Life", path: "/campus-life" },
    { name: "Faculty", path: "/faculty" },
    { name: "Contact Us", path: "/contact" },
  ];

  const scrollToTop = () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  };

  return (
    <footer className="relative bg-gradient-to-b from-[#0b1f33] to-[#071829] text-gray-300">
      {/* MAIN CONTENT */}
      <div className="max-w-7xl mx-auto px-8 py-14 grid grid-cols-1 md:grid-cols-5 gap-12">
        {/* QUICK LINKS */}
        <div className="md:col-span-2">
          <h3 className="text-white text-xl font-semibold mb-6">Quick Links</h3>

          <div className="grid grid-cols-2 gap-x-16 gap-y-3 text-sm">
            {links.map((link) => (
              <Link
                key={link.name}
                to={link.path}
                onClick={scrollToTop}
                className="flex items-center hover:text-white transition"
              >
                <span className="mr-2 text-red-500">»</span>
                {link.name}
              </Link>
            ))}
          </div>
        </div>

        {/* ADDRESS */}
        <div className="md:col-span-3 text-left">
          <h3 className="text-white text-xl font-semibold mb-6 flex items-center">
            <MapPin className="w-5 h-5 mr-2 text-red-500" />
            Authorised Training Centre
          </h3>

          <div className="text-sm leading-relaxed mb-6">
            <p>5th Floor, Computer Seekho Education</p>
            <p>Complex, Gulmohar Road, Juhu</p>
            <p>Scheme, Andheri (W), Mumbai 400 049</p>
            <p>India</p>
          </div>

          <div className="space-y-2 text-sm">
            <p className="flex items-center">
              <Phone className="w-4 h-4 mr-2 text-red-500" />
              9029435311 / 9324095272
            </p>

            <p className="flex items-center">
              <Phone className="w-4 h-4 mr-2 text-red-500" />
              9987062416
            </p>

            <p className="flex items-center">
              <Mail className="w-4 h-4 mr-2 text-red-500" />
              training@computerseekho.com
            </p>
          </div>
        </div>
      </div>

      {/* SOCIAL ICONS */}
      <div className="max-w-7xl mx-auto px-8 pb-6 flex justify-end gap-3">
        <a className="w-10 h-10 bg-gray-700 hover:bg-blue-600 flex items-center justify-center transition">
          <Facebook className="w-5 h-5 text-white" />
        </a>
        <a className="w-10 h-10 bg-gray-700 hover:bg-blue-500 flex items-center justify-center transition">
          <Linkedin className="w-5 h-5 text-white" />
        </a>
        <a className="w-10 h-10 bg-gray-700 hover:bg-red-600 flex items-center justify-center transition">
          <Youtube className="w-5 h-5 text-white" />
        </a>
      </div>

      {/* SCROLL TO TOP */}
      <button
        onClick={scrollToTop}
        className="absolute bottom-6 right-6 w-12 h-12 bg-red-600 hover:bg-red-700 text-white flex items-center justify-center shadow-lg transition"
        aria-label="Scroll to top"
      >
        <ChevronUp className="w-6 h-6" />
      </button>
    </footer>
  );
};

export default Footer;
