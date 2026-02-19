import { Link } from "react-router-dom";

export default function PlacementDropdown({ label, items }) {
  return (
    <div className="relative group">
      {/* PLACEMENT (MATCHES OTHER NAV ITEMS) */}
      <span
        className="
          px-4 py-2
          text-sm font-normal text-[#0A2A66]
          cursor-pointer
          border-b-2 border-transparent
          group-hover:border-[#0A2A66]
        "
      >
        {label}
      </span>

      {/* DROPDOWN */}
      <div
        className="
          absolute left-0 top-full
          hidden
          min-w-[260px]
          bg-white
          shadow-lg
          rounded-md
          group-hover:block
          z-50
        "
      >
        {items.map((item, i) => (
          <Link
            key={i}
            to={item.to}
            className="
              block
              px-4 py-3
              text-sm font-normal text-[#0A2A66]
              border-b-2 border-transparent
              hover:border-[#0A2A66]
            "
          >
            {item.label}
          </Link>
        ))}
      </div>
    </div>
  );
}
