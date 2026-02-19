// import axios from "axios";

// const API_BASE_URL = "http://localhost:8080/api";

// export const getPlacementAlbums = (programCode) =>
//   axios.get(`${API_BASE_URL}/placements/albums`, {
//     params: { programCode },
//   });

// export const getPlacementAlbumImages = (albumId) =>
//   axios.get(`${API_BASE_URL}/placements/albums/${albumId}/images`);



import axios from "axios";

const API_BASE_URL = "http://localhost:8080/api";

export const getPlacementAlbums = (program) =>
  axios.get(`${API_BASE_URL}/placements/albums`, {
    params: { program }, // This matches your backend @RequestParam
  });

export const getPlacementAlbumImages = (albumId) => {
  // ✅ FIXED: Added return statement
  return axios.get(`${API_BASE_URL}/placements/albums/${albumId}/images`);
}
  
