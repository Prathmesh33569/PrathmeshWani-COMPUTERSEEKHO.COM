// import TheLab from "../assets/TheLab.jpg";
// import GuestLecture from "../assets/GuestLecture.jpg";
// import SoftSkills from "../assets/SoftSkills.jpg";
// import MonsoonFootball from "../assets/MonsoonFootball.jpg";
// import LohgadMonsoonTrek from "../assets/LohgadMonsoonTrek.jpg";
// import Library from "../assets/Library.jpg";
// import GroupDiscussion from "../assets/GroupDiscussion.jpg";
// import Coding from "../assets/CodingInFreshAir.jpg";
// import Mock from "../assets/MockInterviews.jpg";
// import Halloween from "../assets/Halloween.jpg";
// // import AtHalloween from "../assets/AtHalloween.jpg";
// import SoftSkillTeamBuilding from "../assets/SoftSkill-TeamBuilding.jpg";
// import SoftSkillTeamBuilding2 from "../assets/SoftSkill-TeamBuilding2.jpg";
// import SharingTheJoyOfXman from "../assets/SharingTheJoyOfX-Max.jpg";
// import DahiHandi from "../assets/DahiHandi.jpg";
// import GovernerOfNagaland from "../assets/WithTheGoverner.jpg";
// import AWalkAroundTheCampus from "../assets/AWalkAroundTheCampus.jpg";

// const campusActivities = [
//   { title: "The Lab", img: TheLab },
//   { title: "Guest Lecture - Tejas Potdar", img: GuestLecture },
//   { title: "Soft Skills - Creativity", img: SoftSkills },
//   { title: "Monsoon Football", img: MonsoonFootball },
//   { title: "Lohgad Monsoon Trek", img: LohgadMonsoonTrek },
//   { title: "The Library", img: Library },
//   { title: "Group Discussion", img: GroupDiscussion },
//   { title: "Coding in Fresh Air", img: Coding },
//   { title: "Mock Interviews", img: Mock },
//   { title: "Halloween", img: Halloween },
//   // { title: "At Halloween", img: AtHalloween },
//   { title: "Soft Skill - Team Building", img: SoftSkillTeamBuilding },
//   { title: "Soft Skill Activity", img: SoftSkillTeamBuilding2 },
//   { title: "Sharing The Joy Of X-Mas", img: SharingTheJoyOfXman },
//   { title: "Dahi Handi", img: DahiHandi },
//   { title: "With Governor of Nagaland", img: GovernerOfNagaland },
//   { title: "A Walk Around Campus", img: AWalkAroundTheCampus },
// ];

// function CampusLife() {
//   return (
//  <div className="w-full bg-white py-12">
//   <div className="max-w-6xl mx-auto px-6">

//     <h1 className="text-3xl font-bold text-center mb-10 text-[#12355b]">
//       Campus Life @ Computer Seekho
//     </h1>

//     <div className="grid grid-cols-1 md:grid-cols-2 gap-6">

//       {campusActivities.map((item, index) => (
//         <div
//           key={index}
//           className="relative overflow-hidden rounded-xl shadow-md"
//         >
//           <img
//             src={item.img}
//             alt={item.title}
//             className="w-full h-[280px] object-cover"
//           />

//           <div className="absolute bottom-0 left-0 w-full
//             bg-gradient-to-t from-red-700/90 via-red-600/60 to-transparent
//             px-6 py-5">

//             <h3 className="text-white text-lg font-semibold">
//               {item.title}
//             </h3>

//           </div>
//         </div>
//       ))}

//     </div>
//   </div>
// </div>

// );

// }

// export default CampusLife;

import { useEffect, useState } from "react";

function CampusLife() {
  const [images, setImages] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    fetch("http://localhost:8080/api/albums/campus-life")
      .then((res) => res.json())
      .then((data) => {
        // We know Campus Life is ONE album
        const campusAlbum = data[0];

        if (campusAlbum && campusAlbum.images) {
          setImages(campusAlbum.images);
        }

        setLoading(false);
      })
      .catch((err) => {
        console.error("Error loading campus life", err);
        setLoading(false);
      });
  }, []);

  if (loading) {
    return <p className="text-center py-10">Loading campus life...</p>;
  }

  return (
    <div className="w-full bg-white py-12">
      <div className="max-w-6xl mx-auto px-6">
        <h1 className="text-3xl font-bold text-center mb-10 text-[#12355b]">
          Campus Life @ Computer Seekho
        </h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {images.map((img, index) => (
            <div
              key={img.imageId}
              className="relative overflow-hidden rounded-xl shadow-md"
            >
              <img
                src={`/${encodeURI(img.imagePath)}`}
                alt="Campus Life"
                className="w-full h-[280px] object-cover"
              />

              <div
                className="absolute bottom-0 left-0 w-full
                bg-gradient-to-t from-red-700/90 via-red-600/60 to-transparent
                px-6 py-5"
              >
                <h3 className="text-white text-lg font-semibold">
                  Campus Life
                </h3>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

export default CampusLife;
