import { useEffect, useState } from "react";
import { useParams } from "react-router-dom";
import { getPlacementAlbumImages } from "../Services/PlacementApi";

const PlacementAlbum = () => {
  const { albumId } = useParams();
  const [images, setImages] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState(null);

  useEffect(() => {
    if (!albumId || isNaN(albumId)) {
      console.error("Invalid albumId:", albumId);
      setError("Invalid album ID");
      setLoading(false);
      return;
    }

    fetchImages();
  }, [albumId]);

  const fetchImages = async () => {
    try {
      setLoading(true);
      setError(null);
      
      const res = await getPlacementAlbumImages(albumId);

      if (res && res.data) {
        setImages(res.data);
      } else {
        setImages([]);
      }
    } catch (err) {
      console.error("Error fetching images:", err);
      setError("Failed to load images");
      setImages([]);
    } finally {
      setLoading(false);
    }
  };

  if (loading) {
    return (
      <div className="p-6 flex justify-center items-center min-h-screen">
        <div className="text-lg">Loading gallery...</div>
      </div>
    );
  }

  if (error) {
    return (
      <div className="p-6">
        <div className="bg-red-100 border border-red-400 text-red-700 px-4 py-3 rounded">
          <p className="font-bold">Error</p>
          <p>{error}</p>
        </div>
      </div>
    );
  }

  // Helper function to extract and format student name from image path
  const extractStudentName = (imagePath) => {
    try {
      // Extract filename from path: /assets/.../ABHISHEK_GUPTA.jpg -> ABHISHEK_GUPTA.jpg
      const filename = imagePath.split('/').pop();
      
      // Remove file extension and _SMVITA suffix
      const nameWithoutExt = filename.replace(/\.(jpg|jpeg|png|gif)$/i, '');
      const nameWithoutSuffix = nameWithoutExt.replace(/_SMVITA$/i, '');
      
      // Convert ABHISHEK_GUPTA to Abhishek Gupta
      const formattedName = nameWithoutSuffix
        .split('_')
        .map(word => word.charAt(0).toUpperCase() + word.slice(1).toLowerCase())
        .join(' ');
      
      return formattedName;
    } catch (error) {
      console.error("Error extracting name:", error);
      return "Student";
    }
  };

  return (
    <div className="p-6">
     <h2 className="text-2xl font-bold mb-6 text-center text-blue-900">Placement Gallery</h2>

      {images.length === 0 ? (
        <p className="text-gray-500">No images found.</p>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6 max-w-6xl mx-auto">
          {images.map((img) => (
            <div 
              key={img.imageId}
              className="bg-white rounded-lg shadow-md hover:shadow-xl transition overflow-hidden"
            >
              <div className="h-80 overflow-hidden">
                <img
                  src={img.imagePath}
                  alt={extractStudentName(img.imagePath)}
                  className="w-full h-full object-cover object-top"
                  onError={(e) => {
                    console.error("Failed to load image:", img.imagePath);
                    e.target.style.display = 'none';
                  }}
                />
              </div>
              <div className="p-4 text-center bg-gray-50">
                <p className="font-semibold text-base text-gray-800">
                  {extractStudentName(img.imagePath)}
                </p>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
};

export default PlacementAlbum;