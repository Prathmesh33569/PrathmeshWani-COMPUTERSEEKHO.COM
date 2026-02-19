import { Outlet } from "react-router-dom";
import Footer from "../Footer/Footer";
import Header from "../Header/Header";

function PublicLayout() {
  return (
    <>
      <Header/>
      <Outlet /> 
      <Footer />
    </>
  );
}

export default PublicLayout;