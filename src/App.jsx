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
import mammoth from "mammoth/mammoth.browser";
import * as pdfjsLib from "pdfjs-dist/legacy/build/pdf.mjs";

pdfjsLib.GlobalWorkerOptions.workerSrc = new URL(
  "pdfjs-dist/legacy/build/pdf.worker.mjs",
  import.meta.url,
).toString();

const INITIAL_FORM = {
  sourceType: "Purple=Sourced",
  unitInterest: "",
  yearsExp: "",
  currentEmp: "",
  hotButtons: "",
  overview: "",
};

const EMAIL_OR_PHONE = /@|\d{3}[\s\-.)]*\d{3}[\s\-.]*\d{4}/;

const cleanValue = (value) =>
  (value || "")
    .replace(/\s+/g, " ")
    .replace(/^[\s\-:|,;.]+|[\s\-:|,;.]+$/g, "")
    .trim();

const titleCase = (value) =>
  value
    .toLowerCase()
    .split(" ")
    .filter(Boolean)
    .map((part) => part[0].toUpperCase() + part.slice(1))
    .join(" ");

const inferSourceType = (text) => {
  const lower = text.toLowerCase();
  if (/\b(applicant|applied|application|inbound)\b/.test(lower)) {
    return "Green=Applicant";
  }
  return "Purple=Sourced";
};

const findByLabel = (text, labels) => {
  for (const rawLine of text.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    for (const label of labels) {
      const pattern = new RegExp(
        `^${label}\\s*[:\\-]\\s*(.+)$`,
        "i",
      );
      const match = line.match(pattern);
      if (match?.[1]) {
        return cleanValue(match[1]);
      }
    }
  }
  return "";
};

const findBestName = (fileName, resumeText, notesText) => {
  const fileNameCandidate = cleanValue(
    fileName
      .replace(/\.[^/.]+$/, "")
      .replace(/resume|cv/gi, " ")
      .replace(/[_-]+/g, " "),
  );
  if (fileNameCandidate) {
    return titleCase(fileNameCandidate);
  }

  const pool = `${resumeText}\n${notesText}`
    .split(/\r?\n/)
    .map((line) => cleanValue(line))
    .filter((line) => line && !EMAIL_OR_PHONE.test(line))
    .slice(0, 12);

  for (const line of pool) {
    if (/^[A-Za-z][A-Za-z\s'.-]{3,40}$/.test(line)) {
      return titleCase(line);
    }
  }

  return "";
};

const QUESTION_LINE = /(^|\s)(q|question|ask|asked|prompt)\s*[:\-]|[?]/i;

const normalizeSentence = (value) =>
  cleanValue(value)
    .replace(/[?]+/g, "")
    .replace(/\s+/g, " ")
    .trim();

const stripQuestionLabel = (value) =>
  value
    .replace(/^\s*(q|question)\s*[:\-]\s*/i, "")
    .replace(/^\s*(a|answer)\s*[:\-]\s*/i, "");

const getCleanFactLines = (text) =>
  (text || "")
    .split(/\r?\n/)
    .map((line) => normalizeSentence(stripQuestionLabel(line)))
    .filter(
      (line) =>
        line.length > 2 &&
        !QUESTION_LINE.test(line) &&
        !EMAIL_OR_PHONE.test(line),
    );

const dedupeLines = (lines) => {
  const seen = new Set();
  return lines.filter((line) => {
    const key = line.toLowerCase();
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
};

const firstMatch = (lines, pattern) => lines.find((line) => pattern.test(line)) || "";

const firstSnippet = (lines, patterns) => {
  for (const pattern of patterns) {
    const hit = firstMatch(lines, pattern);
    if (hit) return hit;
  }
  return "";
};

const formatSentence = (value) => {
  const sentence = normalizeSentence(value);
  if (!sentence) return "";
  return /[.!]$/.test(sentence) ? sentence : `${sentence}.`;
};

const buildDetailedOverview = (notes, resume, extracted) => {
  const lines = dedupeLines(getCleanFactLines(`${notes}\n${resume}`));
  const candidate = extracted.candidateName || "The candidate";
  const yearsExp =
    extracted.yearsExp || firstSnippet(lines, [/\b\d{1,2}\+?\s*(years?|yrs?)\b/i]);
  const specialties =
    extracted.unitInterest ||
    firstSnippet(lines, [
      /\b(med[\s-]?surg|pcu|icu|nicu|picu|er|ed|or|telemetry|step[\s-]?down|labor|delivery|home health)\b/i,
    ]);
  const currentRole = firstSnippet(lines, [
    /\b(currently|presently|serving|working).{0,80}\b(travel|nurse|rn|hospital|facility)\b/i,
  ]);
  const transitionGoal = firstSnippet(lines, [
    /\b(transition|permanent|stable|consisten|long[- ]?term|full[- ]?time)\b/i,
  ]);
  const compensation = firstSnippet(lines, [
    /\$\s?\d{2,3}(?:\s*(?:\/hr|\/hour|per hour))?/i,
    /\b(minimum compensation|pay|rate|salary)\b/i,
  ]);
  const reliability = firstSnippet(lines, [
    /\b(reliable|dependable|attendance|overtime|high standards|passion)\b/i,
  ]);
  const credentials = firstSnippet(lines, [
    /\b(BSN|ADN|MSN|ACLS|BLS|PALS|TNCC|license|licensed|credential)\b/i,
  ]);
  const legalStatus = firstSnippet(lines, [
    /\b(TN visa|green card|sponsorship|work authorization|citizen|permanent resident)\b/i,
  ]);
  const relocation = firstSnippet(lines, [
    /\b(relocat|full[- ]?time|day shift|night shift|not interviewing|competitor)\b/i,
  ]);
  const adminNextStep = firstSnippet(lines, [
    /\b(CGFNS|CES|credential|verification|report|administrative next step|pending)\b/i,
  ]);

  const paragraphs = [];

  const yearsDescriptor = yearsExp ? normalizeSentence(yearsExp) : "";
  const profileParts = [specialties && `specialized focus on ${normalizeSentence(specialties)}`].filter(Boolean);
  if (yearsDescriptor || profileParts.length > 0) {
    paragraphs.push(
      `${candidate} is a veteran nursing professional ${
        yearsDescriptor ? `with ${yearsDescriptor} in clinical practice` : ""
      }${profileParts.length > 0 ? `${yearsDescriptor ? ", including " : "with "}${profileParts.join(", ")}` : ""}.`,
    );
  }

  const objectiveParts = [
    currentRole && normalizeSentence(currentRole),
    transitionGoal && normalizeSentence(transitionGoal),
    compensation && `compensation target: ${normalizeSentence(compensation)}`,
  ].filter(Boolean);
  if (objectiveParts.length > 0) {
    paragraphs.push(`Current placement and goals: ${objectiveParts.join("; ")}.`);
  }

  const qualityParts = [reliability && normalizeSentence(reliability)].filter(Boolean);
  if (qualityParts.length > 0) {
    paragraphs.push(`Professional strengths: ${qualityParts.join("; ")}.`);
  }

  const complianceParts = [
    credentials && normalizeSentence(credentials),
    legalStatus && normalizeSentence(legalStatus),
  ].filter(Boolean);
  if (complianceParts.length > 0) {
    paragraphs.push(`Credentialing and work authorization: ${complianceParts.join("; ")}.`);
  }

  const readinessParts = [
    relocation && normalizeSentence(relocation),
    adminNextStep && `next step: ${normalizeSentence(adminNextStep)}`,
  ].filter(Boolean);
  if (readinessParts.length > 0) {
    paragraphs.push(`Readiness and next actions: ${readinessParts.join("; ")}.`);
  }

  if (paragraphs.length === 0) {
    const fallback = dedupeLines(lines.filter((line) => line.length > 15)).slice(0, 5);
    if (fallback.length === 0) return "";
    return `Candidate summary: ${fallback.map(formatSentence).join(" ")}`;
  }

  return cleanValue(paragraphs.map(formatSentence).join(" "));
};

const buildDetailedHotButtons = (notes, resume, extractedHotButtons) => {
  const lines = dedupeLines(getCleanFactLines(`${notes}\n${resume}`));
  const explicit = normalizeSentence(extractedHotButtons);
  const preferenceLines = lines.filter((line) =>
    /\b(non[- ]?negotiable|dealbreaker|deal breaker|must|requires|cannot|can.?t|won.?t|prefers|preference|avoid|needs|only)\b/i.test(
      line,
    ),
  );
  const availabilityLines = lines.filter((line) =>
    /\b(shift|schedule|weekend|commute|location|distance|pay|rate|salary|benefits|contract)\b/i.test(
      line,
    ),
  );

  const hotList = dedupeLines(
    [explicit, ...preferenceLines.slice(0, 4), ...availabilityLines.slice(0, 2)]
      .map(normalizeSentence)
      .filter(Boolean),
  );

  if (hotList.length === 0) {
    return "No hard dealbreakers captured in the notes. Primary alignment points are compensation, schedule fit, unit match, and location convenience.";
  }

  return `Priority alignment points: ${hotList.map(formatSentence).join(" ")}`;
};

const sanitizeForSubmission = (value) =>
  cleanValue(
    (value || "")
      .split(/\r?\n/)
      .map((line) => normalizeSentence(stripQuestionLabel(line)))
      .filter((line) => line && !QUESTION_LINE.test(line))
      .join(" "),
  );

const extractFromText = (rawNotes, resumeText, fileName) => {
  const notes = rawNotes || "";
  const resume = resumeText || "";
  const combined = `${notes}\n${resume}`;

  const yearsFromLabel = findByLabel(combined, [
    "years of rn experience",
    "rn tenure",
    "years experience",
    "experience",
  ]);
  const yearsFromPattern = combined.match(
    /\b(\d{1,2}\+?\s*(?:years?|yrs?)(?:\s+of)?\s*(?:rn|nursing)?(?:\s+experience)?)\b/i,
  )?.[1];

  const unitInterest =
    findByLabel(combined, [
      "unit interest",
      "position",
      "shift preference",
      "specialty",
      "speciality",
      "department",
    ]) ||
    combined.match(
      /\b(ICU|PICU|NICU|ER|ED|OR|Telemetry|Med[\s-]?Surg|Step[\s-]?Down|L&D|Labor and Delivery)\b/i,
    )?.[1] ||
    "";

  const currentEmp =
    findByLabel(combined, [
      "current employer",
      "employer",
      "current role",
      "currently at",
      "current facility",
    ]) ||
    "";

  const hotButtons =
    findByLabel(combined, [
      "hot buttons",
      "dealbreakers",
      "deal breakers",
      "non-negotiables",
      "must haves",
    ]) ||
    "";

  const rawOverview =
    findByLabel(combined, [
      "sourcing overview",
      "overview",
      "availability",
      "summary",
      "notes",
    ]) ||
    notes;

  const rawExtracted = {
    candidateName: findBestName(fileName, resume, notes),
    sourceType: inferSourceType(combined),
    unitInterest: cleanValue(unitInterest),
    yearsExp: cleanValue(yearsFromLabel || yearsFromPattern || ""),
    currentEmp: cleanValue(currentEmp),
    hotButtons: cleanValue(hotButtons),
    overview: cleanValue(rawOverview),
  };

  return {
    ...rawExtracted,
    hotButtons: buildDetailedHotButtons(notes, resume, rawExtracted.hotButtons),
    overview: buildDetailedOverview(notes, resume, rawExtracted),
  };
};

const mergeFormData = (form, extracted) => ({
  sourceType: form.sourceType || extracted.sourceType,
  unitInterest: form.unitInterest || extracted.unitInterest,
  yearsExp: form.yearsExp || extracted.yearsExp,
  currentEmp: form.currentEmp || extracted.currentEmp,
  hotButtons: form.hotButtons || extracted.hotButtons,
  overview: form.overview || extracted.overview,
});

const normalizeBlock = (value) => {
  const cleaned = sanitizeForSubmission(value);
  return cleaned || "N/A";
};

function App() {
  const [fileStatus, setFileStatus] = useState("(Will extract text from resume if possible)");
  const [candidateName, setCandidateName] = useState("");
  const [rawNotes, setRawNotes] = useState("");
  const [resumeText, setResumeText] = useState("");
  const [uploadedFileName, setUploadedFileName] = useState("");
  const [form, setForm] = useState(INITIAL_FORM);
  const [isExtracting, setIsExtracting] = useState(false);

  const todayFileStamp = useMemo(() => {
    const today = new Date();
    return `${today.getMonth() + 1}-${today.getDate()}`;
  }, []);

  const updateField = (key, value) => {
    setForm((prev) => ({ ...prev, [key]: value }));
  };

  const extractPdfText = async (file) => {
    const bytes = await file.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: bytes }).promise;
    let text = "";

    for (let page = 1; page <= pdf.numPages; page += 1) {
      const pageData = await pdf.getPage(page);
      const content = await pageData.getTextContent();
      text += `${content.items.map((item) => item.str).join(" ")}\n`;
    }

    return text;
  };

  const extractResumeText = async (file) => {
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".txt")) {
      return file.text();
    }
    if (lower.endsWith(".docx")) {
      const arrayBuffer = await file.arrayBuffer();
      const result = await mammoth.extractRawText({ arrayBuffer });
      return result.value || "";
    }
    if (lower.endsWith(".pdf")) {
      return extractPdfText(file);
    }
    return "";
  };

  const applyExtraction = (preferManualValues = true) => {
    const extracted = extractFromText(rawNotes, resumeText, uploadedFileName);
    const mergedForm = preferManualValues
      ? mergeFormData(form, extracted)
      : {
          sourceType: extracted.sourceType,
          unitInterest: extracted.unitInterest,
          yearsExp: extracted.yearsExp,
          currentEmp: extracted.currentEmp,
          hotButtons: extracted.hotButtons,
          overview: extracted.overview,
        };

    setForm(mergedForm);
    if (!candidateName && extracted.candidateName) {
      setCandidateName(extracted.candidateName);
    }

    return { extracted, mergedForm };
  };

  const handleResumeUpload = async (event) => {
    const file = event.target.files?.[0];
    if (!file) return;

    const cleanName = findBestName(file.name, "", "");
    if (cleanName) {
      setCandidateName(cleanName);
    }

    setIsExtracting(true);
    setUploadedFileName(file.name);
    setFileStatus(`Loaded: ${file.name} (extracting text...)`);

    try {
      const extractedResumeText = await extractResumeText(file);
      setResumeText(extractedResumeText);
      setFileStatus(`Loaded: ${file.name} (text extracted)`);
      setTimeout(() => applyExtraction(true), 0);
    } catch (error) {
      setResumeText("");
      setFileStatus(`Loaded: ${file.name} (text extraction not available for this file)`);
    } finally {
      setIsExtracting(false);
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
              children: [
                new TextRun({
                  text: normalizeBlock(value),
                  size: 24,
                }),
              ],
            }),
          ],
        }),
      ],
    });

  const generateDoc = async () => {
    const { extracted, mergedForm } = applyExtraction(true);
    const name = cleanValue(candidateName || extracted.candidateName) || "Candidate";

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
                createRow("Source Type", mergedForm.sourceType),
                createRow("Unit Interest", mergedForm.unitInterest),
                createRow("RN Tenure", mergedForm.yearsExp),
                createRow("Employer/Location", mergedForm.currentEmp),
                createRow("Hot Buttons", mergedForm.hotButtons),
                createRow("Overview/Avail", mergedForm.overview),
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
            <strong>Step 1: Upload Resume</strong>
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
          <textarea
            placeholder="Paste your sourcing/interview notes here..."
            value={rawNotes}
            onChange={(e) => setRawNotes(e.target.value)}
          />
        </div>

        <div className="button-row">
          <button
            type="button"
            className="secondary"
            onClick={() => applyExtraction(true)}
            disabled={isExtracting}
          >
            Auto-Fill From Resume + Notes
          </button>
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

        <button type="button" onClick={generateDoc} disabled={isExtracting}>
          Step 3: Generate TSC Submittal
        </button>
      </div>
    </div>
  );
}

export default App;
