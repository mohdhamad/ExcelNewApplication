
import React, { useState } from "react";
import { ToastContainer, toast } from "react-toastify";
import "react-toastify/dist/ReactToastify.css";
import "./App.css";

function App() {
  const [dumpFile, setDumpFile] = useState(null);
  const [masterFile, setMasterFile] = useState(null);
  const [loading, setLoading] = useState(false);
  const [downloadUrl, setDownloadUrl] = useState(null);

  const handleSubmit = async (e) => {
    e.preventDefault();

    if (!dumpFile || !masterFile) {
      toast.error("Please select both files.");
      return;
    }

    setLoading(true);
    setDownloadUrl(null);

    const formData = new FormData();
    formData.append("dump", dumpFile);
    formData.append("master", masterFile);

    try {
      const response = await fetch("http://localhost:5000/update-excel", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        throw new Error("Server error");
      }

      const blob = await response.blob();
      const url = window.URL.createObjectURL(blob);
      setDownloadUrl(url);
      toast.success("✅ File processed successfully!");
    } catch (error) {
      console.error(error);
      toast.error("❌ Something went wrong. Check server.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="App">
      <h2>🚀 ServiceNow Incident Excel Updater</h2>

      <form onSubmit={handleSubmit}>
        <div>
          <label>Dump Excel File:</label>
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => setDumpFile(e.target.files[0])}
          />
        </div>
        <div>
          <label>Master Excel File:</label>
          <input
            type="file"
            accept=".xlsx"
            onChange={(e) => setMasterFile(e.target.files[0])}
          />
        </div>
        <button type="submit" disabled={loading}>
          {loading ? "Processing..." : "Upload & Process"}
        </button>
      </form>

      {loading && (
        <div className="loader">
          <div className="spinner"></div>
          <p>Processing your files...</p>
        </div>
      )}

      {downloadUrl && (
        <div style={{ marginTop: "20px" }}>
          <a href={downloadUrl} download="updated_master.xlsx">
            📥 Download Updated Master Sheet
          </a>
        </div>
      )}

      <ToastContainer />
    </div>
  );
}

export default App;
