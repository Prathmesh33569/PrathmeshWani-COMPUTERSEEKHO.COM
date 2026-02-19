import { NavLink } from "react-router-dom";
import PlacementDropdown from "../Dropdown/PlacementDropdown";
import CoursesDropdown from "../Dropdown/CoursesDropdown";

function Navbar() {
  return (
    <nav>
      <ul className="max-w-325 mx-auto flex justify-center gap-8 py-4 text-sm font-semibold text-blue-900">
        <NavItem to="/">HOME</NavItem>
        <PlacementDropdown
          label="PLACEMENT"
          style={{ fontWeight: "bold" }}
          items={[
            { label: "Batch Wise Placements", to: "/placement" },
            { label: "Our Recruiters", to: "/more-recruiters" },
          ]}
        />

        <CoursesDropdown />

        <NavItem to="/campus-life">CAMPUS LIFE</NavItem>
        <NavItem to="/faculty">FACULTY</NavItem>
        {/* <NavItem to="/pg-diploma">ONLINE PG DIPLOMA COURSES</NavItem> */}
        <NavItem to="/contact">GET IN TOUCH</NavItem>
      </ul>
    </nav>
  );
}

/* ===== NORMAL NAV ITEM ===== */
function NavItem({ to, children }) {
  return (
    <NavLink
      to={to}
      className={({ isActive }) =>
        `relative cursor-pointer ${
          isActive ? "text-blue-900" : "hover:text-red-600"
        }`
      }
    >
      {({ isActive }) => (
        <>
          {children}
          {isActive && (
            <span className="absolute left-0 -bottom-3 w-full h-0.5 bg-red-600"></span>
          )}
        </>
      )}
    </NavLink>
  );
}

export default Navbar;
