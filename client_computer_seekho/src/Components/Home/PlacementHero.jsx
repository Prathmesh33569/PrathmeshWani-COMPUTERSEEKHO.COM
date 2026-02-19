import { useNavigate } from "react-router-dom";

const PlacementHero = () => {
  const navigate = useNavigate();

  return (
    <section>
      <div className="bg-blue-900 text-white text-center py-3 font-semibold">
        Highest Placement record for PGCP-BDA and PGCP-AC !!
      </div>

      <div className="relative h-[420px]">
        <img
          src="https://images.unsplash.com/photo-1521737604893-d14cc237f11d?w=1600&q=80"
          className="w-full h-full object-cover"
        />

        <button
          onClick={() => navigate("/placement")}
          className="absolute top-1/4 left-1/2 -translate-x-1/2 
                     bg-blue-900 text-white px-6 py-3 rounded 
                     transition-all duration-200 ease-in-out
                     hover:bg-blue-700 hover:scale-105 hover:shadow-lg
                     active:scale-95"
        >
          See Batchwise Placements
        </button>
      </div>
    </section>
  );
};

export default PlacementHero;
