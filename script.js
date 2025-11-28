// ==== CONFIGURATION ====
// Replace this with the URL where you deploy `web_backend.py`
// e.g. "https://your-backend-service.onrender.com"
const BACKEND_BASE_URL = "http://localhost:8000";

// ==== STATE ====
let questions = [];

// ==== DOM HELPERS ====
const $ = (id) => document.getElementById(id);

const subjectEl = $("subject");
const examEl = $("exam");
const sectionEl = $("section");

const qPathEl = $("q-path");
const qAnswerTypeEl = $("q-answer-type");
const qNeedsContextEl = $("q-needs-context");
const qMarksEl = $("q-marks");
const qTextEditorEl = document.getElementById("q-text-editor");
const qMarkEditorEl = document.getElementById("q-mark-editor");

const questionsEmptyEl = $("questions-empty");
const questionsTableBodyEl = $("questions-table").querySelector("tbody");

const addQuestionBtn = $("add-question");
const downloadExcelBtn = $("download-excel");
const clearAllBtn = $("clear-all");
const statusEl = $("status");

// ==== RENDERING ====
function renderQuestions() {
  questionsTableBodyEl.innerHTML = "";

  if (!questions.length) {
    questionsEmptyEl.style.display = "block";
    downloadExcelBtn.disabled = true;
    clearAllBtn.disabled = true;
    return;
  }

  questionsEmptyEl.style.display = "none";
  downloadExcelBtn.disabled = false;
  clearAllBtn.disabled = false;

  questions.forEach((q, idx) => {
    const tr = document.createElement("tr");

    const orderTd = document.createElement("td");
    orderTd.textContent = String(idx + 1);

    const pathTd = document.createElement("td");
    pathTd.textContent = q.path;

    const typeTd = document.createElement("td");
    typeTd.textContent = formatAnswerType(q.answer_type);

    const textTd = document.createElement("td");
    textTd.textContent = q.text_body;

    const marksTd = document.createElement("td");
    marksTd.textContent = q.marks || "";

    const actionsTd = document.createElement("td");
    actionsTd.className = "actions-cell";
    const delBtn = document.createElement("button");
    delBtn.className = "btn";
    delBtn.textContent = "Delete";
    delBtn.addEventListener("click", () => {
      questions.splice(idx, 1);
      renderQuestions();
      updateStatus();
    });
    actionsTd.appendChild(delBtn);

    tr.appendChild(orderTd);
    tr.appendChild(pathTd);
    tr.appendChild(typeTd);
    tr.appendChild(textTd);
    tr.appendChild(marksTd);
    tr.appendChild(actionsTd);

    questionsTableBodyEl.appendChild(tr);
  });
}

function formatAnswerType(value) {
  const v = Number(value);
  if (v === 1) return "1 - Open text answer";
  if (v === 2) return "2 - Multiple choice";
  return "0 - No answer expected";
}

function updateStatus(message) {
  if (message) {
    statusEl.textContent = message;
    return;
  }
  statusEl.textContent = `${questions.length} question(s) in this exam`;
}

// ==== VALIDATION & BUILDERS ====
function validateExamInfo() {
  const subject = subjectEl.value.trim();
  const exam = examEl.value.trim();

  if (!subject) {
    alert("Subject is required.");
    return null;
  }
  if (!exam) {
    alert("Exam code is required.");
    return null;
  }

  return { subject, exam };
}

function handleAddQuestion() {
  const examInfo = validateExamInfo();
  if (!examInfo) return;

  const path = qPathEl.value.trim();
  const questionHtml = getEditorHtml(qTextEditorEl);
  const text = htmlToPlainText(questionHtml);
  const answerType = Number(qAnswerTypeEl.value);
  const markHtml = getEditorHtml(qMarkEditorEl);
  const markScheme = htmlToPlainText(markHtml);
  const marks = Number(qMarksEl.value || 0);
  const needsContext = !!qNeedsContextEl.checked;
  const section = sectionEl.value || "";

  if (!path) {
    alert("Path is required.");
    return;
  }
  if (!text) {
    alert("Question text is required.");
    return;
  }
  if (answerType === 1 || answerType === 2) {
    if (!markScheme) {
      alert("Mark scheme is required for answer types 1 and 2.");
      return;
    }
    if (marks <= 0) {
      alert("Marks must be greater than 0 for answer types 1 and 2.");
      return;
    }
  }

  const order = questions.length + 1;

  questions.push({
    uniqueid: `q-${Date.now()}-${order}`,
    path,
    text_body: text,
    answer_type: answerType,
    mark_scheme: markScheme,
    needs_context: needsContext,
    section,
    topic: "",
    order,
    marks,
  });

  qPathEl.value = "";
  qTextEditorEl.innerHTML = "";
  qMarkEditorEl.innerHTML = "";
  qNeedsContextEl.checked = false;
  qMarksEl.value = "0";
  qPathEl.focus();

  renderQuestions();
  updateStatus("Question added.");
}

async function handleDownloadExcel() {
  const examInfo = validateExamInfo();
  if (!examInfo) return;
  if (!questions.length) {
    alert("Add at least one question before downloading Excel.");
    return;
  }

  const payload = {
    subject: examInfo.subject,
    exam: examInfo.exam,
    questions: questions,
  };

  downloadExcelBtn.disabled = true;
  updateStatus("Generating Excel file...");

  try {
    const res = await fetch(`${BACKEND_BASE_URL}/export_excel`, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!res.ok) {
      let errorMessage = "Failed to generate Excel file.";
      try {
        const data = await res.json();
        if (data && data.error) {
          errorMessage = data.error;
        }
      } catch (_) {
        // ignore JSON parse errors
      }
      alert(errorMessage);
      return;
    }

    const blob = await res.blob();
    const examClean = examInfo.exam.replace(/\s+/g, "_");
    const filename = `${examClean || "questions"}.xlsx`;

    const url = window.URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    window.URL.revokeObjectURL(url);

    updateStatus("Excel downloaded.");
  } catch (err) {
    console.error(err);
    alert(
      `Could not contact backend at ${BACKEND_BASE_URL}.\n\n` +
        "Make sure the Python server is running and the URL is correct."
    );
  } finally {
    downloadExcelBtn.disabled = questions.length === 0;
    updateStatus();
  }
}

function handleClearAll() {
  if (!questions.length) return;
  const confirmed = window.confirm(
    "This will remove all current entries for this exam. Are you sure?"
  );
  if (!confirmed) return;

  questions = [];
  renderQuestions();
  updateStatus("Entries cleared.");
}

// ==== RICH TEXT HELPERS ====
function getEditorHtml(editor) {
  if (!editor) return "";
  return editor.innerHTML.trim();
}

function htmlToPlainText(html) {
  if (!html) return "";
  const temp = document.createElement("div");
  temp.innerHTML = html;
  return (temp.textContent || "").trim();
}

function applyEditorCommand(command) {
  document.execCommand(command, false, null);
}

function insertImageIntoEditor(editorId, file) {
  const editor = document.getElementById(editorId);
  if (!editor || !file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const dataUrl = e.target.result;
    if (!dataUrl) return;

    const img = document.createElement("img");
    img.src = dataUrl;

    // Insert at caret position
    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      editor.appendChild(img);
      return;
    }
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(img);
  };
  reader.readAsDataURL(file);
}

// ==== EVENT WIRING ====
function init() {
  // Wire up rich-text toolbar buttons
  document.querySelectorAll(".editor-toolbar").forEach((toolbar) => {
    const targetId = toolbar.getAttribute("data-target");

    toolbar.querySelectorAll("[data-cmd]").forEach((btn) => {
      const cmd = btn.getAttribute("data-cmd");
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        const editor = document.getElementById(targetId);
        if (!editor) return;
        editor.focus();
        applyEditorCommand(cmd);
      });
    });

    toolbar.querySelectorAll("[data-image-target]").forEach((btn) => {
      const editorId = btn.getAttribute("data-image-target");
      const input = toolbar.querySelector(
        `.image-input[data-editor="${editorId}"]`
      );
      if (!input) return;
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        input.click();
      });
      input.addEventListener("change", () => {
        const file = input.files && input.files[0];
        if (file) {
          insertImageIntoEditor(editorId, file);
        }
        input.value = "";
      });
    });
  });

  addQuestionBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleAddQuestion();
  });

  downloadExcelBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleDownloadExcel();
  });

  clearAllBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleClearAll();
  });

  renderQuestions();
  updateStatus();
}

document.addEventListener("DOMContentLoaded", init);


