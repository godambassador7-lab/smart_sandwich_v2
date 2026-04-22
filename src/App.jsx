import { useMemo, useState } from "react";
import { saveAs } from "file-saver";
import {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

const INITIAL_FORM = {
  sourceType: "Purple=Sourced",
  unitInterest: "",
  yearsExp: "",
  currentEmp: "",
  hotButtons: "",
  overview: "",
};

function App() {
  const [fileStatus, setFileStatus] = useState("(Extracts name & tenure automatically)");
  const [candidateName, setCandidateName] = useState("");
  const [form, setForm] = useState(INITIAL_FORM);

  const todayFileStamp = useMemo(() => {
    const today = new Date();
    return `${today.getMonth() + 1}-${today.getDate()}`;
  }, []);

  const updateField = (key, value) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const handleResumeUpload = (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    setFileStatus(`Loaded: ${file.name}`);
    const cleanName = file.name
      .replace(/\.[^/.]+$/, "")
      .replace(/Resume|CV|-|_/gi, " ")
      .replace(/\s+/g, " ")
      .trim();

    if (cleanName) {
      setCandidateName(cleanName);
    }
  };

  const createRow = (label, value) =>
    new TableRow({
      children: [
        new TableCell({
          width: { size: 30, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              children: [new TextRun({ text: label, bold: true, size: 24 })],
            }),
          ],
        }),
        new TableCell({
          width: { size: 70, type: WidthType.PERCENTAGE },
          children: [
            new Paragraph({
              children: [new TextRun({ text: value || "N/A", size: 24 })],
            }),
          ],
        }),
      ],
    });

  const generateDoc = async () => {
    const name = candidateName || "Candidate";

    const doc = new Document({
      sections: [
        {
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: "TSC Candidate Submittal",
                  bold: true,
                  size: 32,
                  color: "004a99",
                }),
              ],
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                createRow("Name", name),
                createRow("Source Type", form.sourceType),
                createRow("Unit Interest", form.unitInterest),
                createRow("RN Tenure", form.yearsExp),
                createRow("Employer/Location", form.currentEmp),
                createRow("Hot Buttons", form.hotButtons),
                createRow("Overview/Avail", form.overview),
              ],
            }),
          ],
        },
      ],
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, `${name} - TSC Submittal - ${todayFileStamp}.docx`);
  };

  return (
    <div className="page">
      <div className="container">
        <h1>
          Smart Sandwich <span className="emoji">🥪</span>
          <span className="version">v2 High-Speed</span>
        </h1>

        <div className="upload-section">
          <label htmlFor="resumeUpload" className="file-label">
            <strong>📁 Step 1: Upload Resume</strong>
          </label>
          <input
            id="resumeUpload"
            type="file"
            accept=".pdf,.doc,.docx,.txt"
            onChange={handleResumeUpload}
          />
          <p className="file-status">{fileStatus}</p>
        </div>

        <div className="form-group">
          <label>Step 2: Paste Raw Interview Notes Here</label>
          <textarea placeholder="Paste your sourcing/interview notes here..." />
        </div>

        <hr />

        <div className="grid-main">
          <div className="form-group">
            <label>Candidate Name</label>
            <input
              type="text"
              value={candidateName}
              onChange={(e) => setCandidateName(e.target.value)}
            />
          </div>
          <div className="form-group">
            <label>Source Type</label>
            <select
              value={form.sourceType}
              onChange={(e) => updateField("sourceType", e.target.value)}
            >
              <option value="Purple=Sourced">Purple (Sourced)</option>
              <option value="Green=Applicant">Green (Applicant)</option>
            </select>
          </div>
          <div className="form-group">
            <label>Position / Shift Preference</label>
            <input
              type="text"
              value={form.unitInterest}
              onChange={(e) => updateField("unitInterest", e.target.value)}
            />
          </div>
          <div className="form-group">
            <label>Years of RN Experience</label>
            <input
              type="text"
              value={form.yearsExp}
              onChange={(e) => updateField("yearsExp", e.target.value)}
            />
          </div>
          <div className="form-group">
            <label>Current Employer / Location</label>
            <input
              type="text"
              value={form.currentEmp}
              onChange={(e) => updateField("currentEmp", e.target.value)}
            />
          </div>
          <div className="form-group">
            <label>Hot Buttons (Dealbreakers)</label>
            <input
              type="text"
              value={form.hotButtons}
              onChange={(e) => updateField("hotButtons", e.target.value)}
            />
          </div>
          <div className="form-group full-width">
            <label>Sourcing Overview & Availability</label>
            <textarea
              value={form.overview}
              onChange={(e) => updateField("overview", e.target.value)}
            />
          </div>
        </div>

        <button type="button" onClick={generateDoc}>
          Step 3: Generate TSC Submittal
        </button>
      </div>
    </div>
  );
}

export default App;
