import Hero from "../Components/Home/Hero";
import Courses from "../Components/Home/Courses";
import WhyChoose from "../Components/Home/WhyChoose";
import PlacementHero from "../Components/Home/PlacementHero";
import Recruiters from "../Components/Home/Recruiters";
import OPDC from "./OPDC"; // ✅ correct import
import CourseOffered from "../Components/Home/CourseOffered";

const Home = () => {
  return (
    <>
      <Hero />
      <CourseOffered></CourseOffered>
      <WhyChoose />
      <PlacementHero />
      <Recruiters />
    </>
  );
};

export default Home;
