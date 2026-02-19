import { useEffect, useState } from "react";
import { useNavigate } from "react-router-dom";
import { getPlacementAlbums } from "../Services/PlacementApi";

function Placement() {
  const navigate = useNavigate();
  const [dacAlbums, setDacAlbums] = useState([]);
  const [dbdaAlbums, setDbdaAlbums] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    fetchAlbums();
  }, []);

  const fetchAlbums = async () => {
    try {
      setLoading(true);
      setError(null);

      const [dacRes, dbdaRes] = await Promise.all([
        getPlacementAlbums("DAC"),
        getPlacementAlbums("DBDA"),
      ]);

      console.log("DAC Response:", dacRes.data);
      console.log("DBDA Response:", dbdaRes.data);

      setDacAlbums(dacRes.data || []);
      setDbdaAlbums(dbdaRes.data || []);
    } catch (err) {
      console.error("Error fetching albums:", err);
      setError(err.message || "Failed to load placement albums");
    } finally {
      setLoading(false);
    }
  };

  const renderAlbums = (albums) => {
    if (albums.length === 0) {
      return <p className="text-gray-500">No placements available yet.</p>;
    }

    return (
      <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
        {albums.map((album) => {
          const coverSrc = album.coverImagePath?.trim() || null;

          return (
            <div
              key={album.albumId}
              onClick={() => navigate(`/placement/${album.albumId}`)}
              className="cursor-pointer bg-white rounded-xl shadow hover:shadow-lg transition overflow-hidden"
            >
              {coverSrc ? (
                <img
                  src={coverSrc}
                  alt={album.albumName}
                  className="h-70 w-full object-cover"
                  onError={(e) => {
                    console.error("Failed to load image:", coverSrc);
                    e.target.style.display = 'none';
                    if (!e.target.nextElementSibling) {
                      const fallback = document.createElement('div');
                      fallback.className = 'h-64 w-full bg-gradient-to-br from-blue-500 to-purple-600 flex items-center justify-center';
                      fallback.innerHTML = `<span class="text-white text-xl font-bold">${album.albumName}</span>`;
                      e.target.parentElement.insertBefore(fallback, e.target);
                    }
                  }}
                />
              ) : (
                <div className="h-70 w-full bg-linear-to-br from-blue-500 to-purple-600 flex items-center justify-center">
                  <span className="text-white text-xl font-bold">
                    {album.albumName}
                  </span>
                </div>
              )}

              <div className="p-4">
                <h3 className="font-semibold text-lg">{album.albumName}</h3>
                <p className="text-sm text-gray-600">
                  {album.albumDescription}
                </p>
              </div>
            </div>
          );
        })}
      </div>
    );
  };

  if (loading) {
    return (
      <div className="p-6 flex justify-center items-center min-h-screen">
        <div className="text-lg">Loading placements...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-6">
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded">
          <p className="font-bold">Error</p>
          <p>{error}</p>
          <button
            onClick={fetchAlbums}
            className="mt-2 bg-red-600 text-white px-4 py-2 rounded hover:bg-red-700"
          >
            Retry
          </button>
        </div>
      </div>
    );
  }

  return (
    <div className="p-6 space-y-12">
      <section>
        <h2 className="text-2xl font-bold mb-4 text-center text-blue-900">PG DAC Placements</h2>
        {renderAlbums(dacAlbums)}
      </section>

      <section>
       <h2 className="text-2xl font-bold mb-4 text-center text-blue-900">PG DBDA Placements</h2>
        {renderAlbums(dbdaAlbums)}
      </section>
    </div>
  );
}

export default Placement;



 {/* <img
            src={album.coverImage || "/placeholder.jpg"}
            alt={album.albumName}
            className="h-48 w-full object-cover rounded-t-xl"
          /> */}