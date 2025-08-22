import React, { useState, useEffect } from "react";
import { useAuth } from "../utils/auth";
import { useNavigate } from "react-router-dom";
import Dashboard from "../components/Dashboard";
import axios from "axios";

function GradePreview() {
  const { logout } = useAuth();
  const navigate = useNavigate();

  // Original GradePreview state
  const [students, setStudents] = useState([]);
  const [file, setFile] = useState(null);
  const [fileName, setFileName] = useState(""); // Added to track file name
  const [topper, setTopper] = useState(null);
  const [activeDept, setActiveDept] = useState("TECHNICAL");
  const [isUploading, setIsUploading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);

  // New state for tabs and history
  const [activeTab, setActiveTab] = useState("create"); // "create" or "history"

  // Generated Files state (from GeneratedFiles component)
  const [generatedFiles, setGeneratedFiles] = useState([]);
  const [loadingFiles, setLoadingFiles] = useState(true);
  const [selectedFile, setSelectedFile] = useState(null);
  const [selectedFileStudents, setSelectedFileStudents] = useState([]);
  const [isEditingHistory, setIsEditingHistory] = useState(false);
  const [historyActiveDept, setHistoryActiveDept] = useState("TECHNICAL");
  const [isUpdating, setIsUpdating] = useState(false);
  const [isDownloading, setIsDownloading] = useState(false);

  // Add state for popup
  const [showFilesPopup, setShowFilesPopup] = useState(false);

  // Add state for custom date (only for Excel generation, not saved to database)
  const [customDate, setCustomDate] = useState("");

  const API_BASE = "http://localhost:5000";

  const handleUpload = async () => {
    if (!file) {
      alert("Please select a file first.");
      return;
    }

    setIsUploading(true);
    const form = new FormData();
    form.append("file", file);

    try {
      const res = await axios.post(
        "http://localhost:5000/upload/trainee",
        form
      );

      const dataWithGrades = res.data.students.map((student) => {
        const dept = student.department?.trim().toUpperCase();
        const gradeCount =
          dept === "PRODUCTION" || dept === "Production" ? 18 : 15;

        return {
          ...student,
          ...Object.fromEntries(
            Array.from({ length: gradeCount }, (_, i) => [`${i + 1}G`, ""])
          ),
        };
      });

      setStudents(dataWithGrades);
      setFileName(file.name); // Store the file name
    } catch (err) {
      console.error(err);
      alert("Upload failed. Please try again.");
    } finally {
      setIsUploading(false);
    }
  };

  const handleRemoveFile = () => {
    setFile(null);
    setFileName("");
    setStudents([]);
    setTopper(null);
  };

  const handleGradeChange = (index, field, value) => {
    const updated = [...students];

    // For performance field (over_all), allow 1 decimal place
    if (field === "over_all") {
      // Allow up to 1 decimal place
      if (value === "" || /^\d*\.?\d{0,1}$/.test(value)) {
        updated[index][field] = value;
      }
    } else {
      // For other fields, only allow whole numbers
      if (value === "" || /^\d+$/.test(value)) {
        updated[index][field] = value;
      }
    }

    setStudents(updated);

    // Calculate top performer based on overall score (over_all field)
    const studentsWithGrades = updated.filter(
      (s) => s.over_all && Number(s.over_all) > 0
    );
    if (studentsWithGrades.length > 0) {
      const top = studentsWithGrades.reduce((prev, current) =>
        Number(current.over_all) > Number(prev.over_all) ? current : prev
      );
      setTopper(top);
    } else {
      setTopper(null);
    }
  };

  // History-specific grade change handler
  const handleHistoryGradeChange = (index, field, value) => {
    const updated = [...selectedFileStudents];

    // For performance field (over_all), allow 1 decimal place
    if (field === "over_all") {
      // Allow up to 1 decimal place
      if (value === "" || /^\d*\.?\d{0,1}$/.test(value)) {
        // Convert empty string to null for database consistency
        updated[index][field] = value === "" ? null : value;
      }
    } else {
      // For other fields, only allow whole numbers
      if (value === "" || /^\d+$/.test(value)) {
        // Convert empty string to null for database consistency
        updated[index][field] = value === "" ? null : value;
      }
    }

    setSelectedFileStudents(updated);
  };

  const departments = ["TECHNICAL", "PRODUCTION", "SUPPORT"];

  const filteredStudents = students.filter((s) => {
    const dept = s.department?.trim().toUpperCase() || "";
    if (activeDept === "TECHNICAL") {
      return dept === "TECHNICAL" || dept === "IT";
    }
    if (activeDept === "PRODUCTION") {
      return dept === "PRODUCTION" || dept === "PROD";
    }
    if (activeDept === "SUPPORT") {
      return (
        dept !== "TECHNICAL" &&
        dept !== "IT" &&
        dept !== "PRODUCTION" &&
        dept !== "PROD"
      );
    }
    return false;
  });

  // Filtered students for history view
  const filteredHistoryStudents = selectedFileStudents.filter((s) => {
    const dept = s.department?.trim().toUpperCase() || "";
    if (historyActiveDept === "TECHNICAL") {
      return dept === "TECHNICAL" || dept === "IT";
    }
    if (historyActiveDept === "PRODUCTION") {
      return dept === "PRODUCTION" || dept === "PROD";
    }
    if (historyActiveDept === "SUPPORT") {
      return (
        dept !== "TECHNICAL" &&
        dept !== "IT" &&
        dept !== "PRODUCTION" &&
        dept !== "PROD"
      );
    }
    return false;
  });

  // Grading layouts
  const gradingStructures = {
    PRODUCTION: {
      groups: [
        { label: "NTOP", fields: ["WI", "CO", "5S", "BO", "CBO", "SDG"] },
        { label: "WVS", fields: ["OHSA", "WE", "UJC", "ISO", "PO", "HR"] },
        { label: "EQUIP", fields: ["WI2", "ELEX", "CM", "SPC"] },
        { label: "ASSESSMENT", fields: ["PROD", "DS"] },
      ],
    },
    SUPPORT: {
      groups: [
        { label: "NTOP", fields: ["WI", "CO", "5S", "BO", "CBO", "SDG"] },
        { label: "WVS", fields: ["OHSA", "WE", "UJC", "ISO", "PO", "HR"] },
        { label: "EQUIP", fields: ["PerDev"] },
        { label: "ASSESSMENT", fields: ["Supp", "DS"] },
      ],
    },
    TECHNICAL: {
      groups: [
        { label: "NTOP", fields: ["WI", "CO", "5S", "BO", "CBO", "SDG"] },
        { label: "WVS", fields: ["OHSA", "WE", "UJC", "ISO", "PO", "HR"] },
        { label: "EQUIP", fields: ["AppDev"] },
        { label: "ASSESSMENT", fields: ["Tech", "DS"] },
      ],
    },
  };

  const handleGenerateExcel = async () => {
    if (students.length === 0) {
      alert("No student data available. Please upload a file first.");
      return;
    }

    setIsGenerating(true);
    try {
      // Use the original file name for the generated report
      const baseFileName = fileName.replace(/\.[^/.]+$/, ""); // Remove extension
      const reportData = {
        students,
        originalFileName: baseFileName,
      };

      const res = await axios.post(
        "http://localhost:5000/api/generate/excel",
        reportData,
        { responseType: "blob" }
      );

      // Extract filename from Content-Disposition safely
      const disposition = res.headers["content-disposition"];
      let filename = `${baseFileName}-REPORT.xlsx`; // Use original file name
      if (disposition) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match && match[1]) {
          filename = match[1];
        }
      }

      const blob = new Blob([res.data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      // Refresh files list after generating new file
      fetchGeneratedFiles();
    } catch (err) {
      console.error(err);
      alert("Download failed. Please try again.");
    } finally {
      setIsGenerating(false);
    }
  };

  // Generate Excel with custom date for history files
  const handleGenerateHistoryExcel = async () => {
    if (!selectedFile || selectedFileStudents.length === 0) {
      alert("No student data available.");
      return;
    }

    setIsGenerating(true);
    try {
      // Prepare students data with custom date if provided
      const studentsWithCustomDate = selectedFileStudents.map((student) => ({
        ...student,
        // Use custom date if provided, otherwise keep original
        date_of_immersion: customDate || student.date_of_immersion,
      }));

      const reportData = {
        students: studentsWithCustomDate,
        originalFileName: selectedFile.filename.replace(/\.[^/.]+$/, ""),
      };

      const res = await axios.post(
        "http://localhost:5000/api/generate/excel",
        reportData,
        { responseType: "blob" }
      );

      // Extract filename from Content-Disposition safely
      const disposition = res.headers["content-disposition"];
      let filename = `${selectedFile.filename.replace(
        /\.[^/.]+$/,
        ""
      )}-UPDATED-REPORT.xlsx`;
      if (disposition) {
        const match = disposition.match(/filename="?([^"]+)"?/);
        if (match && match[1]) {
          filename = match[1];
        }
      }

      const blob = new Blob([res.data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = window.URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      window.URL.revokeObjectURL(url);

      alert("Excel report generated successfully with updated data!");
    } catch (err) {
      console.error(err);
      alert("Download failed. Please try again.");
    } finally {
      setIsGenerating(false);
    }
  };

  // Generated Files API functions
  const fetchGeneratedFiles = async () => {
    setLoadingFiles(true);
    try {
      const response = await fetch(`${API_BASE}/api/generated-files`);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      const data = await response.json();
      setGeneratedFiles(data.files || []);
    } catch (error) {
      console.error("Error fetching files:", error);
      alert("Failed to fetch files: " + error.message);
    }
    setLoadingFiles(false);
  };

  const fetchFileDetails = async (fileId) => {
    try {
      const response = await fetch(`${API_BASE}/api/generated-files/${fileId}`);
      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }
      const data = await response.json();

      setSelectedFileStudents(data.students || []);
      setSelectedFile(data.file);
      // Reset custom date when selecting a new file
      setCustomDate("");
      // Close popup after selecting a file
      setShowFilesPopup(false);
    } catch (error) {
      console.error("Error fetching file details:", error);
      alert("Failed to fetch file details: " + error.message);
    }
  };

  const updateFile = async () => {
    if (!selectedFile) return;

    setIsUpdating(true);
    try {
      const response = await fetch(
        `${API_BASE}/api/generated-files/${selectedFile.id}`,
        {
          method: "PUT",
          headers: {
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            students: selectedFileStudents,
            batch: selectedFile.batch,
            school: selectedFile.school,
            date_of_immersion: selectedFile.date_of_immersion,
          }),
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      const data = await response.json();
      console.log("File updated successfully:", data);
      alert("File updated successfully!");
      setIsEditingHistory(false);

      // Refresh the files list to show updated info
      fetchGeneratedFiles();
    } catch (error) {
      console.error("Error updating file:", error);
      alert("Failed to update file: " + error.message);
    }
    setIsUpdating(false);
  };

  const deleteFile = async (fileId) => {
    if (!confirm("Are you sure you want to delete this file?")) return;

    try {
      const response = await fetch(
        `${API_BASE}/api/generated-files/${fileId}`,
        {
          method: "DELETE",
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // Remove from local state
      setGeneratedFiles(generatedFiles.filter((f) => f.id !== fileId));

      // Clear selection if deleted file was selected
      if (selectedFile && selectedFile.id === fileId) {
        setSelectedFile(null);
        setSelectedFileStudents([]);
        setIsEditingHistory(false);
      }

      alert("File deleted successfully!");
    } catch (error) {
      console.error("Error deleting file:", error);
      alert("Failed to delete file: " + error.message);
    }
  };

  const downloadFile = async (fileId, filename) => {
    setIsDownloading(true);
    try {
      const response = await fetch(
        `${API_BASE}/api/generated-files/${fileId}/download`,
        {
          method: "GET",
        }
      );

      if (!response.ok) {
        throw new Error(`HTTP error! status: ${response.status}`);
      }

      // Get the blob from response
      const blob = await response.blob();

      // Create download link
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = url;
      link.download = filename;
      document.body.appendChild(link);
      link.click();

      // Cleanup
      window.URL.revokeObjectURL(url);
      document.body.removeChild(link);

      console.log("Download completed for:", filename);
    } catch (error) {
      console.error("Error downloading file:", error);
      alert("Failed to download file: " + error.message);
    }
    setIsDownloading(false);
  };

  // Load generated files on component mount
  useEffect(() => {
    fetchGeneratedFiles();
  }, []);

  // Clear selected file when switching tabs but preserve uploaded file data
  useEffect(() => {
    if (activeTab === "create") {
      setSelectedFile(null);
      setSelectedFileStudents([]);
      setIsEditingHistory(false);
      setCustomDate("");
      setShowFilesPopup(false);
      // Don't clear file, fileName, students, or topper when switching back to create tab
    }
  }, [activeTab]);

  return (
    <div className="relative bg-neutral-900 flex min-h-screen text-white overflow-hidden font-arial">
      {/* Background */}
      <div className="fixed inset-0 bg-cover bg-center blur-lg -z-10"></div>
      <div className="absolute inset-0 bg-black/10 -z-10"></div>

      {/* Add custom styles to hide number input arrows and set small input size */}
      <style>
        {`
          input[type="number"]::-webkit-outer-spin-button,
          input[type="number"]::-webkit-inner-spin-button {
            -webkit-appearance: none;
            margin: 0;
          }
          
          input[type="number"] {
            -moz-appearance: textfield;
          }
          
          .grade-input {
            width: 50px !important;
            min-width: 50px;
            max-width: 50px;
          }
        `}
      </style>

      {/* Files Popup Overlay */}
      {showFilesPopup && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-start justify-center pt-20">
          <div className="bg-zinc-800/95 backdrop-blur-sm p-6 rounded-xl shadow-2xl max-w-6xl w-full mx-4 max-h-[80vh] overflow-hidden">
            <div className="flex justify-between items-center mb-6">
              <h2 className="text-2xl font-bold text-white">
                📚 Generated Files
              </h2>
              <button
                onClick={() => setShowFilesPopup(false)}
                className="text-white hover:text-red-400 text-2xl font-bold p-2"
              >
                ✕
              </button>
            </div>

            <div className="mb-4">
              <button
                onClick={fetchGeneratedFiles}
                className="px-4 py-2 bg-gradient-to-r from-purple-600 to-purple-700 hover:from-purple-700 hover:to-purple-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg text-sm"
                disabled={loadingFiles}
              >
                {loadingFiles ? "🔄 Loading..." : "🔄 Refresh Files"}
              </button>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4 max-h-[60vh] overflow-y-auto pr-2">
              {generatedFiles.length === 0 ? (
                <div className="col-span-full text-center text-gray-400 py-12">
                  {loadingFiles ? (
                    <div className="animate-pulse">
                      <div className="text-4xl mb-4">⏳</div>
                      <div className="text-sm">Loading files...</div>
                    </div>
                  ) : (
                    <div>
                      <div className="text-4xl mb-4">📂</div>
                      <div className="text-sm">No files found</div>
                    </div>
                  )}
                </div>
              ) : (
                generatedFiles.map((file) => (
                  <div
                    key={file.id}
                    className={`p-4 rounded-xl border cursor-pointer transition-all duration-300 shadow-md hover:shadow-lg ${
                      selectedFile?.id === file.id
                        ? "bg-gradient-to-r from-purple-600 to-purple-700 border-purple-400 transform scale-[1.02]"
                        : "bg-zinc-600/80 border-zinc-500 hover:bg-zinc-600"
                    }`}
                    onClick={() => {
                      fetchFileDetails(file.id);
                      setActiveTab("history");
                    }}
                  >
                    <div
                      className="font-semibold text-base text-white truncate mb-2"
                      title={file.batch}
                    >
                      {file.batch}
                    </div>
                    <div
                      className="text-sm text-gray-300 truncate mb-1"
                      title={file.school}
                    >
                      🏫 {file.school}
                    </div>
                    <div className="text-sm text-gray-300 mb-1">
                      👥 {file.student_count || 0} students
                    </div>
                    <div className="text-sm text-gray-300 mb-1">
                      📅{" "}
                      {file.created_at
                        ? new Date(file.created_at).toLocaleDateString()
                        : "N/A"}
                    </div>
                    {file.average_performance && (
                      <div className="text-sm text-green-400 font-semibold">
                        📊 Avg:{" "}
                        {parseFloat(file.average_performance).toFixed(1)}%
                      </div>
                    )}
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}

      <Dashboard />

      {/* Main Content on the right */}
      <div className="flex-1 p-4 mt-16 mr-8 overflow-y-auto">
        <div className="w-full max-w-[1920px] mx-auto">
          {/* Tab Navigation */}
          <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl mb-8">
            <div className="flex gap-4">
              <button
                onClick={() => setActiveTab("create")}
                className={`px-8 py-4 rounded-xl font-bold text-lg transition-all duration-300 shadow-lg ${
                  activeTab === "create"
                    ? "bg-gradient-to-r from-[#9d4edd] to-[#c77dff] text-white shadow-[0px_0px_20px_rgba(157,78,221,0.6)] transform scale-105"
                    : "bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white hover:from-[#7b1fa2] hover:to-[#9c27b0] hover:scale-105"
                }`}
              >
                📝 Create New Report
              </button>
              <button
                onClick={() => {
                  setActiveTab("history");
                  setShowFilesPopup(true);
                }}
                className={`px-8 py-4 rounded-xl font-bold text-lg transition-all duration-300 shadow-lg relative ${
                  activeTab === "history"
                    ? "bg-gradient-to-r from-[#9d4edd] to-[#c77dff] text-white shadow-[0px_0px_20px_rgba(157,78,221,0.6)] transform scale-105"
                    : "bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white hover:from-[#7b1fa2] hover:to-[#9c27b0] hover:scale-105"
                }`}
              >
                📚 Generated Files History
                {generatedFiles.length > 0 && (
                  <span className="absolute -top-2 -right-2 bg-gradient-to-r from-yellow-500 to-amber-500 text-black text-sm font-bold px-3 py-1 rounded-full min-w-[28px] flex items-center justify-center shadow-lg">
                    {generatedFiles.length}
                  </span>
                )}
              </button>
            </div>
          </div>

          {/* Create New Report Tab */}
          {activeTab === "create" && (
            <>
              {/* File Upload Section */}
              <div className="bg-zinc-700/90 backdrop-blur-sm p-8 rounded-xl shadow-2xl mb-8">
                <h2 className="text-2xl font-bold mb-6 text-white">
                  📁 Upload Trainee Excel File
                </h2>

                {!file ? (
                  <div className="flex flex-col gap-4">
                    <input
                      type="file"
                      onChange={(e) => {
                        setFile(e.target.files[0]);
                        setFileName(e.target.files[0]?.name || "");
                      }}
                      className="bg-zinc-600/80 text-white p-4 rounded-xl file:bg-gradient-to-r file:from-violet-600 file:to-purple-600 file:text-white file:border-0 file:rounded-lg file:px-4 file:py-2 file:mr-4 file:font-semibold file:cursor-pointer file:hover:from-violet-700 file:hover:to-purple-700 file:transition-all"
                      accept=".xlsx,.xls"
                    />
                    <button
                      className="w-fit rounded-xl bg-gradient-to-r from-[#3737ff] to-[#6a5acd] px-8 py-4 text-white font-bold hover:from-[#4169e1] hover:to-[#7b68ee] transition-all duration-300 shadow-lg disabled:opacity-50 disabled:cursor-not-allowed"
                      onClick={handleUpload}
                      disabled={isUploading || !file}
                    >
                      {isUploading ? "🔄 Uploading..." : "⬆️ Upload File"}
                    </button>
                  </div>
                ) : (
                  <div className="bg-zinc-600/80 p-6 rounded-xl">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-4">
                        <span className="text-2xl">📄</span>
                        <div>
                          <p className="text-white font-semibold text-lg">
                            {fileName}
                          </p>
                          <p className="text-gray-300 text-sm">
                            {students.length > 0
                              ? `${students.length} students loaded`
                              : "Ready to upload"}
                          </p>
                        </div>
                      </div>
                      <div className="flex gap-3">
                        {students.length === 0 && (
                          <button
                            className="rounded-xl bg-gradient-to-r from-[#3737ff] to-[#6a5acd] px-6 py-3 text-white font-bold hover:from-[#4169e1] hover:to-[#7b68ee] transition-all duration-300 shadow-lg disabled:opacity-50 disabled:cursor-not-allowed"
                            onClick={handleUpload}
                            disabled={isUploading}
                          >
                            {isUploading ? "🔄 Uploading..." : "⬆️ Upload"}
                          </button>
                        )}
                        <button
                          className="rounded-xl bg-gradient-to-r from-red-600 to-red-700 px-6 py-3 text-white font-bold hover:from-red-700 hover:to-red-800 transition-all duration-300 shadow-lg"
                          onClick={handleRemoveFile}
                        >
                          🗑️ Remove
                        </button>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              {/* Top Performer Section */}
              {topper && (
                <div className="bg-gradient-to-r from-zinc-700/90 to-zinc-600/90 backdrop-blur-sm p-8 rounded-xl shadow-2xl mb-8">
                  <h2 className="text-2xl font-bold mb-6 text-white">
                    🏆 Top Performer
                  </h2>
                  <div className="text-center bg-gradient-to-br from-zinc-800/80 to-zinc-900/80 p-8 rounded-xl border border-[#9d4edd] shadow-2xl">
                    <div className="flex items-center justify-center space-x-6 mb-4">
                      <span className="text-6xl animate-bounce">🏆</span>
                      <div className="text-left">
                        <div className="font-bold text-3xl text-yellow-400 mb-2">
                          {topper.last_name}, {topper.first_name}{" "}
                          {topper.middle_name || ""}
                        </div>
                        <div className="text-xl text-white mb-2">
                          <span className="text-yellow-300 font-semibold">
                            Overall Score:
                          </span>
                          <span className="text-white font-bold ml-3 text-2xl bg-gradient-to-r from-green-400 to-green-600 bg-clip-text">
                            {topper.over_all || "N/A"}
                          </span>
                        </div>
                        <div className="text-lg text-white">
                          <span className="text-purple-300 font-semibold">
                            Department:
                          </span>
                          <span className="text-white font-semibold ml-2">
                            {topper.department}
                          </span>
                          <span className="text-purple-300 font-semibold ml-4">
                            Strand:
                          </span>
                          <span className="text-white font-semibold ml-2">
                            {topper.strand}
                          </span>
                        </div>
                      </div>
                    </div>
                    <div className="bg-gradient-to-r from-yellow-500 to-amber-500 text-black px-6 py-3 rounded-full text-lg font-bold">
                      🌟 Current Top Performer 🌟
                    </div>
                  </div>
                </div>
              )}

              {/* Department Selection */}
              {students.length > 0 && (
                <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl mb-8">
                  <h3 className="text-xl font-bold text-white mb-4">
                    Select Department
                  </h3>
                  <div className="flex gap-4">
                    {departments.map((dept) => (
                      <button
                        key={dept}
                        onClick={() => setActiveDept(dept)}
                        className={`px-6 py-3 rounded-xl font-bold text-lg transition-all duration-300 shadow-lg ${
                          activeDept === dept
                            ? "bg-gradient-to-r from-[#9d4edd] to-[#c77dff] text-white shadow-[0px_0px_15px_rgba(157,78,221,0.6)] transform scale-105"
                            : "bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white hover:from-[#7b1fa2] hover:to-[#9c27b0] hover:scale-105"
                        }`}
                      >
                        {dept}
                      </button>
                    ))}
                  </div>
                </div>
              )}

              {/* Grades Table */}
              {filteredStudents.length > 0 && (
                <div className="bg-zinc-700/90 backdrop-blur-sm p-8 rounded-xl shadow-2xl mb-8">
                  <h2 className="text-2xl font-bold mb-6 text-white">
                    📊 {activeDept} Department Grades
                  </h2>
                  <div className="overflow-x-auto rounded-xl">
                    <table className="w-full border-collapse text-sm bg-zinc-800/50 rounded-xl overflow-hidden">
                      <thead>
                        <tr>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Name
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Strand
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Department
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            School
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Batch
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Date
                          </th>
                          <th
                            className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                            rowSpan="2"
                          >
                            Performance Appraisal
                          </th>
                          {gradingStructures[activeDept].groups.map(
                            (group, idx) => (
                              <th
                                key={idx}
                                className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                colSpan={group.fields.length}
                              >
                                {group.label}
                              </th>
                            )
                          )}
                        </tr>
                        <tr>
                          {gradingStructures[activeDept].groups.flatMap(
                            (group) =>
                              group.fields.map((field) => (
                                <th
                                  key={field}
                                  className="border border-zinc-600 p-3 text-center bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                >
                                  {field}
                                </th>
                              ))
                          )}
                        </tr>
                      </thead>

                      <tbody>
                        {filteredStudents.map((s, i) => (
                          <tr
                            key={i}
                            className="text-center hover:bg-zinc-600/30 transition-colors"
                          >
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80 font-medium">
                              {s.last_name}, {s.first_name} {s.middle_name}
                            </td>
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                              {s.strand}
                            </td>
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                              {s.department}
                            </td>
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                              {s.school}
                            </td>
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                              {s.batch}
                            </td>
                            <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                              {s.date_of_immersion}
                            </td>

                            <td className="border border-zinc-600 p-3 bg-zinc-800/80">
                              <input
                                type="number"
                                value={s.over_all || ""}
                                onChange={(e) =>
                                  handleGradeChange(
                                    students.indexOf(s),
                                    "over_all",
                                    e.target.value
                                  )
                                }
                                className="grade-input p-2 bg-zinc-700 border border-zinc-500 rounded-lg text-white text-center focus:border-[#6a0dad] focus:outline-none focus:ring-2 focus:ring-[#6a0dad]/20"
                                min="0"
                                max="100"
                                step="0.1"
                                placeholder="0"
                              />
                            </td>

                            {gradingStructures[activeDept].groups.flatMap(
                              (group) =>
                                group.fields.map((field) => (
                                  <td
                                    key={field}
                                    className="border border-zinc-600 p-3 bg-zinc-800/80"
                                  >
                                    <input
                                      type="number"
                                      value={s[field] || ""}
                                      onChange={(e) =>
                                        handleGradeChange(
                                          students.indexOf(s),
                                          field,
                                          e.target.value
                                        )
                                      }
                                      className="grade-input p-2 bg-zinc-700 border border-zinc-500 rounded-lg text-white text-center focus:border-[#6a0dad] focus:outline-none focus:ring-2 focus:ring-[#6a0dad]/20"
                                      min="0"
                                      max="100"
                                      step="1"
                                      placeholder="0"
                                    />
                                  </td>
                                ))
                            )}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Generate Button */}
              {students.length > 0 && (
                <div className="bg-zinc-700/90 backdrop-blur-sm p-8 rounded-xl shadow-2xl mb-8 text-center">
                  <button
                    className="rounded-xl bg-gradient-to-r from-[#a361ef] to-[#c77dff] px-10 py-4 text-white font-bold text-xl hover:from-[#9551df] hover:to-[#b666f0] transition-all duration-300 shadow-2xl disabled:opacity-50 disabled:cursor-not-allowed transform hover:scale-105"
                    onClick={handleGenerateExcel}
                    disabled={isGenerating}
                  >
                    {isGenerating
                      ? "⏳ Generating Report..."
                      : "📊 Generate Excel Report"}
                  </button>
                </div>
              )}
            </>
          )}

          {/* History Tab */}
          {activeTab === "history" && (
            <div className="w-full">
              {!selectedFile ? (
                <div className="flex items-center justify-center h-[600px] bg-zinc-700/90 backdrop-blur-sm rounded-xl shadow-2xl">
                  <div className="text-center">
                    <div className="text-8xl mb-6">📁</div>
                    <h2 className="text-3xl font-bold mb-4 text-white">
                      Select a File
                    </h2>
                    <p className="text-xl text-gray-300 mb-6">
                      Click "Generated Files History" to choose a file to view
                      and edit
                    </p>
                    <button
                      onClick={() => setShowFilesPopup(true)}
                      className="px-8 py-4 bg-gradient-to-r from-purple-600 to-purple-700 hover:from-purple-700 hover:to-purple-800 rounded-xl transition-all duration-300 font-bold text-white shadow-lg"
                    >
                      📚 Open Files List
                    </button>
                  </div>
                </div>
              ) : (
                <>
                  {/* File Header */}
                  <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl mb-6">
                    <div className="flex items-start justify-between mb-4">
                      <div className="flex-1 min-w-0">
                        <h2
                          className="text-2xl font-bold text-white mb-3 truncate"
                          title={selectedFile.filename}
                        >
                          📄 {selectedFile.filename}
                        </h2>
                        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 text-sm text-gray-300">
                          <div className="flex items-center gap-2">
                            <span>🏫</span>
                            <span
                              className="truncate"
                              title={selectedFile.school}
                            >
                              {selectedFile.school}
                            </span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span>👥</span>
                            <span>{selectedFile.batch}</span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span>📅</span>
                            <span>
                              {selectedFile.date_of_immersion
                                ? new Date(
                                    selectedFile.date_of_immersion
                                  ).toLocaleDateString()
                                : "N/A"}
                            </span>
                          </div>
                          <div className="flex items-center gap-2">
                            <span>🎓</span>
                            <span>{selectedFileStudents.length} students</span>
                          </div>
                        </div>
                      </div>

                      <div className="flex flex-wrap gap-2 ml-4">
                        <button
                          onClick={() => setShowFilesPopup(true)}
                          className="px-4 py-2 bg-gradient-to-r from-purple-600 to-purple-700 hover:from-purple-700 hover:to-purple-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg text-sm"
                        >
                          📚 Change File
                        </button>
                        {!isEditingHistory ? (
                          <>
                            <button
                              onClick={() => setIsEditingHistory(true)}
                              className="px-4 py-2 bg-gradient-to-r from-blue-600 to-blue-700 hover:from-blue-700 hover:to-blue-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg text-sm"
                            >
                              ✏️ Edit
                            </button>
                            <button
                              onClick={() =>
                                downloadFile(
                                  selectedFile.id,
                                  selectedFile.filename
                                )
                              }
                              disabled={isDownloading}
                              className="px-4 py-2 bg-gradient-to-r from-green-600 to-green-700 hover:from-green-700 hover:to-green-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg disabled:opacity-50 text-sm"
                            >
                              {isDownloading ? "⏳..." : "⬇️ Download"}
                            </button>
                            <button
                              onClick={() => deleteFile(selectedFile.id)}
                              className="px-4 py-2 bg-gradient-to-r from-red-600 to-red-700 hover:from-red-700 hover:to-red-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg text-sm"
                            >
                              🗑️ Delete
                            </button>
                          </>
                        ) : (
                          <>
                            <button
                              onClick={updateFile}
                              disabled={isUpdating}
                              className="px-4 py-2 bg-gradient-to-r from-purple-600 to-purple-700 hover:from-purple-700 hover:to-purple-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg disabled:opacity-50 text-sm"
                            >
                              {isUpdating ? "⏳ Saving..." : "💾 Save Changes"}
                            </button>
                            <button
                              onClick={() => {
                                setIsEditingHistory(false);
                                fetchFileDetails(selectedFile.id);
                              }}
                              className="px-4 py-2 bg-gradient-to-r from-gray-600 to-gray-700 hover:from-gray-700 hover:to-gray-800 rounded-lg transition-all duration-300 font-semibold text-white shadow-lg text-sm"
                            >
                              ❌ Cancel
                            </button>
                          </>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* Custom Date Input and Generate Excel Button - Only in Edit Mode */}
                  {isEditingHistory && (
                    <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl mb-6">
                      <h3 className="text-lg font-bold text-white mb-4">
                        📊 Generate Updated Excel Report
                      </h3>
                      <div className="flex flex-col sm:flex-row gap-4 items-end">
                        <div className="flex-1">
                          <label className="block text-white font-semibold mb-2">
                            📅 Custom Date for Excel (Optional)
                          </label>
                          <input
                            type="date"
                            value={customDate}
                            onChange={(e) => setCustomDate(e.target.value)}
                            className="w-full p-3 bg-zinc-600 border border-zinc-500 rounded-lg text-white focus:border-purple-500 focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                            placeholder="Select date for Excel report"
                          />
                          <p className="text-gray-400 text-sm mt-1">
                            Leave empty to use original dates from the file
                          </p>
                        </div>
                        <button
                          onClick={handleGenerateHistoryExcel}
                          disabled={isGenerating}
                          className="px-6 py-3 bg-gradient-to-r from-[#a361ef] to-[#c77dff] hover:from-[#9551df] hover:to-[#b666f0] rounded-lg transition-all duration-300 font-bold text-white shadow-lg disabled:opacity-50 transform hover:scale-105 whitespace-nowrap"
                        >
                          {isGenerating
                            ? "⏳ Generating..."
                            : "📊 Generate Excel"}
                        </button>
                      </div>
                    </div>
                  )}

                  {/* Department Tabs for History */}
                  {isEditingHistory && (
                    <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl mb-6">
                      <h3 className="text-lg font-bold text-white mb-4">
                        Select Department to Edit
                      </h3>
                      <div className="flex flex-wrap gap-3">
                        {departments.map((dept) => {
                          const deptStudentCount = selectedFileStudents.filter(
                            (s) => {
                              const studentDept =
                                s.department?.trim().toUpperCase() || "";
                              if (dept === "TECHNICAL") {
                                return (
                                  studentDept === "TECHNICAL" ||
                                  studentDept === "IT"
                                );
                              }
                              if (dept === "PRODUCTION") {
                                return (
                                  studentDept === "PRODUCTION" ||
                                  studentDept === "PROD"
                                );
                              }
                              if (dept === "SUPPORT") {
                                return (
                                  studentDept !== "TECHNICAL" &&
                                  studentDept !== "IT" &&
                                  studentDept !== "PRODUCTION" &&
                                  studentDept !== "PROD"
                                );
                              }
                              return false;
                            }
                          ).length;

                          return (
                            <button
                              key={dept}
                              onClick={() => setHistoryActiveDept(dept)}
                              className={`px-5 py-3 rounded-xl font-semibold text-base transition-all duration-300 shadow-lg ${
                                historyActiveDept === dept
                                  ? "bg-gradient-to-r from-[#9d4edd] to-[#c77dff] text-white shadow-[0px_0px_15px_rgba(157,78,221,0.6)] transform scale-105"
                                  : "bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white hover:from-[#7b1fa2] hover:to-[#9c27b0] hover:scale-105"
                              }`}
                            >
                              {dept} ({deptStudentCount})
                            </button>
                          );
                        })}
                      </div>
                    </div>
                  )}

                  {/* History Students Content */}
                  {isEditingHistory ? (
                    <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl">
                      <h3 className="text-xl font-bold mb-4 text-white">
                        📊 {historyActiveDept} Department Grades (Edit Mode) -{" "}
                        {filteredHistoryStudents.length} students
                      </h3>

                      {filteredHistoryStudents.length === 0 ? (
                        <div className="text-center text-gray-400 py-12">
                          <div className="text-6xl mb-4">👥</div>
                          <div className="text-lg">
                            No students found in {historyActiveDept} department
                          </div>
                        </div>
                      ) : (
                        <div className="overflow-x-auto rounded-xl">
                          <table className="w-full border-collapse text-sm bg-zinc-800/50 rounded-xl overflow-hidden">
                            <thead>
                              <tr>
                                <th
                                  className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                  rowSpan="2"
                                >
                                  Name
                                </th>
                                <th
                                  className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                  rowSpan="2"
                                >
                                  Strand
                                </th>
                                <th
                                  className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                  rowSpan="2"
                                >
                                  Department
                                </th>
                                <th
                                  className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                  rowSpan="2"
                                >
                                  Performance
                                </th>
                                {gradingStructures[
                                  historyActiveDept
                                ].groups.map((group, idx) => (
                                  <th
                                    key={idx}
                                    className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                    colSpan={group.fields.length}
                                  >
                                    {group.label}
                                  </th>
                                ))}
                              </tr>
                              <tr>
                                {gradingStructures[
                                  historyActiveDept
                                ].groups.flatMap((group) =>
                                  group.fields.map((field) => (
                                    <th
                                      key={field}
                                      className="border border-zinc-600 p-3 bg-gradient-to-r from-[#6a0dad] to-[#8a2be2] text-white font-bold"
                                    >
                                      {field}
                                    </th>
                                  ))
                                )}
                              </tr>
                            </thead>

                            <tbody>
                              {filteredHistoryStudents.map((student, index) => (
                                <tr
                                  key={student.id || index}
                                  className="hover:bg-zinc-600/30 transition-colors"
                                >
                                  <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80 font-medium">
                                    {student.last_name}, {student.first_name}{" "}
                                    {student.middle_name}
                                  </td>
                                  <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                                    {student.strand}
                                  </td>
                                  <td className="border border-zinc-600 p-3 text-white bg-zinc-800/80">
                                    {student.department}
                                  </td>
                                  <td className="border border-zinc-600 p-3 bg-zinc-800/80">
                                    <input
                                      type="number"
                                      value={student.over_all || ""}
                                      onChange={(e) =>
                                        handleHistoryGradeChange(
                                          selectedFileStudents.indexOf(student),
                                          "over_all",
                                          e.target.value
                                        )
                                      }
                                      className="grade-input p-2 bg-zinc-700 border border-zinc-500 rounded-lg text-white text-center focus:border-purple-500 focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                                      min="0"
                                      max="100"
                                      step="0.1"
                                      placeholder="0"
                                    />
                                  </td>

                                  {gradingStructures[
                                    historyActiveDept
                                  ].groups.flatMap((group) =>
                                    group.fields.map((field) => (
                                      <td
                                        key={field}
                                        className="border border-zinc-600 p-3 bg-zinc-800/80"
                                      >
                                        <input
                                          type="number"
                                          value={student[field] || ""}
                                          onChange={(e) =>
                                            handleHistoryGradeChange(
                                              selectedFileStudents.indexOf(
                                                student
                                              ),
                                              field,
                                              e.target.value
                                            )
                                          }
                                          className="grade-input p-2 bg-zinc-700 border border-zinc-500 rounded-lg text-white text-center focus:border-purple-500 focus:outline-none focus:ring-2 focus:ring-purple-500/20"
                                          min="0"
                                          max="100"
                                          step="1"
                                          placeholder="0"
                                        />
                                      </td>
                                    ))
                                  )}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      )}
                    </div>
                  ) : (
                    // View Mode for History - Better card layout
                    <div className="bg-zinc-700/90 backdrop-blur-sm p-6 rounded-xl shadow-2xl">
                      <h3 className="text-xl font-bold mb-4 text-white">
                        👥 Students Overview - {selectedFileStudents.length}{" "}
                        total students
                      </h3>

                      {selectedFileStudents.length === 0 ? (
                        <div className="text-center text-gray-400 py-12">
                          <div className="text-6xl mb-4">📊</div>
                          <div className="text-lg">
                            No student data available
                          </div>
                        </div>
                      ) : (
                        <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 2xl:grid-cols-4 gap-4">
                          {selectedFileStudents.map((student, index) => (
                            <div
                              key={student.id || index}
                              className="bg-zinc-600/80 p-5 rounded-xl shadow-lg hover:shadow-xl transition-all duration-300 border border-zinc-500 hover:border-zinc-400"
                            >
                              <div
                                className="font-semibold text-lg text-white mb-3 truncate"
                                title={`${student.last_name}, ${student.first_name}`}
                              >
                                {student.last_name}, {student.first_name}
                              </div>
                              <div className="text-white space-y-2 text-sm">
                                <div className="flex items-center gap-2">
                                  <span>📚</span>
                                  <span className="truncate">
                                    {student.strand}
                                  </span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <span>🏢</span>
                                  <span className="truncate">
                                    {student.department}
                                  </span>
                                </div>
                                <div className="flex items-center gap-2">
                                  <span>📊</span>
                                  <span>Performance:</span>
                                  <span className="font-bold text-green-400 text-base">
                                    {student.over_all || "N/A"}
                                  </span>
                                </div>
                              </div>
                            </div>
                          ))}
                        </div>
                      )}
                    </div>
                  )}
                </>
              )}
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default GradePreview;
