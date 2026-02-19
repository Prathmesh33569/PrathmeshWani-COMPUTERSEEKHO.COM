import Navbar from "../Navbar/NavBar";

import appLogo3 from "../../assets/CSlogo3.jpg";

function Header() {
  return (
    <header className="w-full bg-white">
      <div className="max-w-[1300px] mx-auto flex items-center justify-center gap-8 px-4 py-4">
        {/* <div className="w-[420px] flex justify-center">
          <img
            src={appLogo1}
            alt="VTA Logo"
            className="h-24 w-auto object-contain"
          />
        </div> */}

        <div className="w-[820px] flex justify-center">
          <img
            src={appLogo3}
            alt="SM Vita Logo"
            className="h-40  object-contain"
          />
        </div>

        {/* <div className="w-[420px] flex justify-center">
          <img
            src={appLogo2}
            alt="CDAC ACTS"
            className="h-24 w-auto object-contain"
          />
        </div> */}
      </div>

      <Navbar></Navbar>
    </header>
  );
}

export default Header;
