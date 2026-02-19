// import api from "./api";

// const batchService = {
//   createBatch: async (batchData, batchLogo) => {
//     const formData = new FormData();

//     // Append text fields
//     formData.append("batchName", batchData.batchName);
//     formData.append("batchStartDate", batchData.batchStartDate || "");
//     formData.append("batchEndDate", batchData.batchEndDate || "");
//     formData.append("courseId", batchData.courseId || "");
//     formData.append("presentationDate", batchData.presentationDate || "");
//     formData.append(
//       "finalPresentationDate",
//       batchData.finalPresentationDate || "",
//     );
//     formData.append("courseFees", batchData.courseFees || "");
//     formData.append("courseFeesFrom", batchData.courseFeesFrom || "");
//     formData.append("courseFeesTo", batchData.courseFeesTo || "");
//     formData.append("batchIsActive", batchData.batchIsActive ?? true);

//     // Append image
//     if (batchLogo) {
//       formData.append("batchLogo", batchLogo); // must match @RequestParam name
//     }

//     const response = await api.post("/batches", formData, {
//       headers: {
//         "Content-Type": "multipart/form-data",
//       },
//     });

//     return response.data;
//   },

//   updateBatch: async (batchId, batchData, batchLogo) => {
//     const formData = new FormData();

//     // Append text fields
//     formData.append("batchName", batchData.batchName);
//     formData.append("batchStartDate", batchData.batchStartDate || "");
//     formData.append("batchEndDate", batchData.batchEndDate || "");
//     formData.append("courseId", batchData.courseId || "");
//     formData.append("presentationDate", batchData.presentationDate || "");
//     formData.append(
//       "finalPresentationDate",
//       batchData.finalPresentationDate || "",
//     );
//     formData.append("courseFees", batchData.courseFees || "");
//     formData.append("batchIsActive", batchData.batchIsActive ?? true);

//     // Append image if provided (new image uploaded)
//     if (batchLogo) {
//       formData.append("batchLogo", batchLogo); // must match @RequestParam name
//     }

//     const response = await api.put(`/batches/${batchId}`, formData, {
//       headers: {
//         "Content-Type": "multipart/form-data",
//       },
//     });

//     return response.data;
//   },
// };

// export default batchService;




import api from "./api";

const batchService = {

  createBatch: async (batchData) => {
    const response = await api.post("/batches", batchData);
    return response.data;
  },

  updateBatch: async (batchId, batchData) => {
    const response = await api.put(`/batches/${batchId}`, batchData);
    return response.data;
  },

};

export default batchService;
