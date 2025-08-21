import React, { useState } from "react";
import { useAuth } from "../utils/auth";
import { useNavigate } from "react-router-dom";
import Dashboard from "../components/Dashboard";
import axios from "axios";

function GradePreview() {
  const { logout } = useAuth();
  const navigate = useNavigate();

  const [students, setStudents] = useState([]);
  const [file, setFile] = useState(null);
  const [topper, setTopper] = useState(null);
  const [activeDept, setActiveDept] = useState("TECHNICAL");
  const [isUploading, setIsUploading] = useState(false);
  const [isGenerating, setIsGenerating] = useState(false);

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
    } catch (err) {
      console.error(err);
      alert("Upload failed. Please try again.");
    } finally {
      setIsUploading(false);
    }
  };

  const handleGradeChange = (index, field, value) => {
    const updated = [...students];
    updated[index][field] = value;
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
      const res = await axios.post(
        "http://localhost:5000/api/generate/excel",
        { students },
        { responseType: "blob" }
      );

      // Extract filename from Content-Disposition safely
      const disposition = res.headers["content-disposition"];
      let filename = "IMMERSION-GENERATED.xlsx"; // fallback
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
    } catch (err) {
      console.error(err);
      alert("Download failed. Please try again.");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <div className="relative bg-neutral-900 flex min-h-screen text-white overflow-hidden font-arial">
      {/* Background */}
      <div className="fixed inset-0 bg-cover bg-center blur-lg -z-10"></div>
      {/* Optional dark overlay */}
      <div className="absolute inset-0 bg-black/10 -z-10"></div>

      {/* Add custom styles to hide number input arrows */}
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
        `}
      </style>

      <Dashboard />

      {/* Main Content on the right */}
      <div className="flex-1 p-3 mt-13 mr-9 overflow-y-auto">
        <div className="w-[95%] mx-auto">
          {/* Combined File Upload, Department Selection, and Top Performer Section */}
          <div className="w-full max-w-[1800px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
            {/* File Upload Section */}
            <div className="mb-6">
              <h2 className="text-xl font-bold mb-[15px]">
                Upload Trainee Excel File
              </h2>
              <div className="flex items-center gap-4">
                <input
                  type="file"
                  onChange={(e) => setFile(e.target.files[0])}
                  className="bg-zinc-600 text-white p-2 rounded-md file:bg-violet-600 file:text-white file:border-0 file:rounded file:px-3 file:py-1 file:mr-3"
                  accept=".xlsx,.xls"
                />
                <button
                  className="rounded-md bg-[#3737ff7e] px-6 py-2 hover:bg-[#3737ff] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                  onClick={handleUpload}
                  disabled={isUploading || !file}
                >
                  {isUploading ? "Uploading..." : "Upload"}
                </button>
              </div>
            </div>

            {/* Department Selection (moved to top-left, no title) */}

            {/* Top Performer */}
            {topper && (
              <div>
                <h2 className="text-xl font-bold mb-[15px]">Top Performer</h2>
                <div className="text-center text-[1.2rem] bg-[#222] p-4 rounded-lg border border-[#9d4edd]">
                  <div className="flex items-center justify-center space-x-4 mb-3">
                    <span className="text-4xl animate-bounce">🏆</span>
                    <div className="text-left">
                      <div className="font-bold text-2xl text-yellow-400 mb-1">
                        {topper.last_name}, {topper.first_name}{" "}
                        {topper.middle_name || ""}
                      </div>
                      <div className="text-lg text-gray-300">
                        <span className="text-yellow-300">Overall Score:</span>
                        <span className="text-white font-bold ml-2 text-xl">
                          {topper.over_all || "N/A"}
                        </span>
                      </div>
                      <div className="text-md text-gray-400 mt-1">
                        <span className="text-purple-300">Department:</span>
                        <span className="text-white font-semibold ml-2">
                          {topper.department}
                        </span>{" "}
                        |<span className="text-purple-300 ml-2">Strand:</span>
                        <span className="text-white font-semibold ml-2">
                          {topper.strand}
                        </span>
                      </div>
                    </div>
                  </div>
                  <div className="bg-gradient-to-r from-yellow-600 to-yellow-500 text-black px-4 py-2 rounded-full text-sm font-bold">
                    🌟 Current Top Performer 🌟
                  </div>
                </div>
              </div>
            )}
          </div>

          {students.length > 0 && (
            <div className="mb-6 flex items-center gap-4">
              {/* Department buttons on the left */}
              <div className="flex gap-2">
                {departments.map((dept) => (
                  <button
                    key={dept}
                    onClick={() => setActiveDept(dept)}
                    className={`px-4 py-2 rounded-lg font-bold transition-all duration-300 ${
                      activeDept === dept
                        ? "bg-[#9d4edd] text-white shadow-[0px_0px_10px_rgba(157,78,221,0.6)]"
                        : "bg-[#6a0dad] text-white hover:bg-[#8a2be2]"
                    }`}
                  >
                    {dept}
                  </button>
                ))}
              </div>
            </div>
          )}

          {/* Grades Table with integrated CSS styles */}
          {filteredStudents.length > 0 && (
            <div className="w-full max-w-[1800px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
              <h2 className="text-xl font-bold mb-[15px]">
                {activeDept} Department Grades
              </h2>
              {/* Table container with integrated styles */}
              <div className="w-full overflow-x-auto">
                <table className="w-full border-collapse text-[0.85rem]">
                  <thead>
                    {/* First row: group headers */}
                    <tr>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Name
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Strand
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Department
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        School
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Batch
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Date
                      </th>
                      <th
                        className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                        rowSpan="2"
                      >
                        Performance Appraisal
                      </th>
                      {gradingStructures[activeDept].groups.map(
                        (group, idx) => (
                          <th
                            key={idx}
                            className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
                            colSpan={group.fields.length}
                          >
                            {group.label}
                          </th>
                        )
                      )}
                    </tr>
                    {/* Second row: field headers */}
                    <tr>
                      {gradingStructures[activeDept].groups.flatMap((group) =>
                        group.fields.map((field) => (
                          <th
                            key={field}
                            className="border border-[#333] p-1.5 text-center bg-[#6a0dad] text-white"
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
                        className="text-center hover:bg-zinc-700/30 transition-colors"
                      >
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.last_name}, {s.first_name} {s.middle_name}
                        </td>
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.strand}
                        </td>
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.department}
                        </td>
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.school}
                        </td>
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.batch}
                        </td>
                        <td className="border border-[#333] p-1.5 text-center text-white bg-zinc-800">
                          {s.date_of_immersion}
                        </td>

                        {/* Performance Appraisal input */}
                        <td className="border border-[#333] p-1.5 text-center bg-zinc-800">
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
                            className="w-[90%] p-1 bg-[#222] border border-[#444] rounded text-white text-center focus:border-[#6a0dad] focus:outline-none"
                            min="0"
                            max="100"
                          />
                        </td>

                        {/* Dynamic fields with integrated input styles */}
                        {gradingStructures[activeDept].groups.flatMap((group) =>
                          group.fields.map((field) => (
                            <td
                              key={field}
                              className="border border-[#333] p-1.5 text-center bg-zinc-800"
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
                                className="w-[90%] p-1 bg-[#222] border border-[#444] rounded text-white text-center focus:border-[#6a0dad] focus:outline-none"
                                min="0"
                                max="100"
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
            <div className="w-full max-w-[1800px] bg-zinc-500 p-6 rounded-md mx-auto mb-8">
              <button
                className="rounded-md bg-[#a361ef] px-6 py-3 text-white font-semibold hover:bg-[#9551df] transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
                onClick={handleGenerateExcel}
                disabled={isGenerating}
              >
                {isGenerating ? "Generating..." : "Generate Excel Report"}
              </button>
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

export default GradePreview;
