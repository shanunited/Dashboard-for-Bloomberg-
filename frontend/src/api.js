import axios from "axios";

const api = axios.create({
  baseURL: import.meta.env.VITE_API_BASE_URL || "http://127.0.0.1:8000",
});

export async function compareFiles({ yesterdayFile, todayFile }) {
  const formData = new FormData();
  formData.append("yesterday_file", yesterdayFile);
  formData.append("today_file", todayFile);

  const response = await api.post("/compare", formData, {
    headers: { "Content-Type": "multipart/form-data" },
  });

  return response.data;
}

export async function generateReport({ reportFile, reportDate }) {
  const formData = new FormData();
  formData.append("report_file", reportFile);
  formData.append("report_date", reportDate);

  const response = await api.post("/generate-report", formData, {
    headers: { "Content-Type": "multipart/form-data" },
    responseType: "blob",
  });

  return response.data;
}

export async function generateCombinedReport({ stockFile, indexFile }) {
  const formData = new FormData();
  formData.append("stock_file", stockFile);
  formData.append("index_file", indexFile);

  const response = await api.post("/generate-combined-report", formData, {
    headers: { "Content-Type": "multipart/form-data" },
    responseType: "blob",
  });

  return response.data;
}

export async function processIndexRp({ indexRpFile }) {
  const formData = new FormData();
  formData.append("index_rp_file", indexRpFile);

  const response = await api.post("/index-rp", formData, {
    headers: { "Content-Type": "multipart/form-data" },
  });

  return response.data;
}
