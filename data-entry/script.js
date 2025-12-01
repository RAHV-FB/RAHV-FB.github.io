// ==== STATE ====
const state = {
  exam: {
    subject: "",
    exam: "",
  },
  sets: [], // {id, label, section}
  questions: [], // questions with full metadata + rich text HTML
  currentSetId: null,
  editingQuestionId: null,
};

// ==== DOM HELPERS ====
const $ = (id) => document.getElementById(id);

// Basic elements
const subjectEl = $("subject");
const examEl = $("exam");

// Sets
const setsListEl = $("sets-list");
const newSetBtn = $("btn-new-set");
const editSetBtn = $("btn-edit-set");
const deleteSetBtn = $("btn-delete-set");
const setSelectEl = $("q-set-select");

// Question form
const qFormTitleEl = $("question-form-title");
const qPathEl = $("q-path");
const qAnswerTypeEl = $("q-answer-type");
const qNeedsContextEl = $("q-needs-context");
const qMarksEl = $("q-marks");
const qTextEditorEl = $("q-text-editor");
const qMarkEditorEl = $("q-mark-editor");

const addBtn = $("btn-add-question");
const addContinueBtn = $("btn-add-continue");
const updateBtn = $("btn-update-question");
const cancelEditBtn = $("btn-cancel-edit");

// Preview / export
const previewContainerEl = $("preview-container");
const previewEmptyEl = $("preview-empty");
const exportExcelBtn = $("btn-export-excel");
const clearAllBtn = $("btn-clear-all");
const statusEl = $("status");

// ==== UTILITIES ====
function uuid() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, (c) => {
    const r = (Math.random() * 16) | 0;
    const v = c === "x" ? r : (r & 0x3) | 0x8;
    return v.toString(16);
  });
}

function formatAnswerTypeLabel(v) {
  const value = Number(v);
  if (value === 1) return "1 - Open text answer";
  if (value === 2) return "2 - Multiple choice";
  return "0 - No answer expected (context only)";
}

function htmlToPlainText(html) {
  if (!html) return "";
  const temp = document.createElement("div");
  temp.innerHTML = html;
  return (temp.textContent || "").trim();
}

// ==== EXAM VALIDATION ====
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
  // Updated: must start with M or N, then any characters
  if (!/^[MN]/.test(exam)) {
    alert("Exam code must start with M or N (e.g., M22, N23, M22A, etc.).");
    return null;
  }

  state.exam.subject = subject;
  state.exam.exam = exam;
  return { subject, exam };
}

// ==== SETS ====
function refreshSetSelect() {
  setSelectEl.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "(Select set)";
  setSelectEl.appendChild(placeholder);

  state.sets.forEach((s) => {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.section ? `${s.label} (Section ${s.section})` : s.label;
    setSelectEl.appendChild(opt);
  });

  if (state.currentSetId) {
    setSelectEl.value = state.currentSetId;
  }
}

function renderSetsList() {
  setsListEl.innerHTML = "";
  state.sets.forEach((s) => {
    const li = document.createElement("li");
    li.dataset.id = s.id;
    li.textContent = s.section ? `${s.label} (Section ${s.section})` : s.label;
    if (s.id === state.currentSetId) li.classList.add("selected");
    li.addEventListener("click", () => {
      state.currentSetId = s.id;
      renderSetsList();
      refreshSetSelect();
    });
    setsListEl.appendChild(li);
  });
}

function ensureSetSelected() {
  if (!state.currentSetId || !state.sets.find((s) => s.id === state.currentSetId)) {
    alert("Please create and select a question set first.");
    return false;
  }
  return true;
}

function handleNewSet() {
  const label = window.prompt("Set label (e.g., Question 10):", "");
  if (!label || !label.trim()) return;
  const section = window.prompt("Section (optional, e.g., A or B):", "") || "";
  const set = { id: uuid(), label: label.trim(), section: section.trim() };
  state.sets.push(set);
  state.currentSetId = set.id;
  renderSetsList();
  refreshSetSelect();
  renderPreview();
}

function handleEditSet() {
  if (!state.currentSetId) {
    alert("Select a set to edit.");
    return;
  }
  const set = state.sets.find((s) => s.id === state.currentSetId);
  if (!set) return;
  const newLabel = window.prompt("Set label:", set.label) ?? "";
  if (!newLabel.trim()) {
    alert("Set label is required.");
    return;
  }
  const newSection = window.prompt("Section (optional, e.g., A or B):", set.section || "") ?? "";
  set.label = newLabel.trim();
  set.section = newSection.trim();

  // Update questions belonging to this set
  state.questions.forEach((q) => {
    if (q.set_id === set.id) {
      q.set_label = set.label;
      q.section = set.section;
    }
  });

  renderSetsList();
  refreshSetSelect();
  renderPreview();
}

function handleDeleteSet() {
  if (!state.currentSetId) {
    alert("Select a set to delete.");
    return;
  }
  const set = state.sets.find((s) => s.id === state.currentSetId);
  if (!set) return;
  const ok = window.confirm(
    `Delete set "${set.label}" and all its questions? This cannot be undone.`
  );
  if (!ok) return;

  state.sets = state.sets.filter((s) => s.id !== set.id);
  state.questions = state.questions.filter((q) => q.set_id !== set.id);
  state.currentSetId = state.sets[0]?.id || null;

  renderSetsList();
  refreshSetSelect();
  renderPreview();
  updateStatus();
}

// ==== RICH TEXT HELPERS ====
function getEditorHtml(editor) {
  if (!editor) return "";
  return editor.innerHTML.trim();
}

function applyEditorCommand(command, editor) {
  if (!editor) return;
  editor.focus();
  if (command === "bold" || command === "italic") {
    document.execCommand(command, false, null);
  } else if (command === "monospace") {
    document.execCommand("fontName", false, "Courier New");
  } else if (command === "normal") {
    document.execCommand("removeFormat", false, null);
  }
}

// Fixed image insertion: always inserts into the correct editor at cursor position
function insertImageIntoEditor(editorId, file) {
  const editor = document.getElementById(editorId);
  if (!editor || !file) return;

  // Ensure editor is focused
  editor.focus();

  const reader = new FileReader();
  reader.onload = (e) => {
    const dataUrl = e.target.result;
    if (!dataUrl) return;

    // Create img element for inline display
    const img = document.createElement("img");
    img.src = dataUrl;
    img.style.maxWidth = "100%";
    img.style.height = "auto";
    img.style.display = "block";
    img.style.margin = "0.5rem 0";
    img.style.border = "1px solid #ddd";
    img.style.borderRadius = "4px";

    // Get selection - ensure it's within our editor
    const selection = window.getSelection();
    let range;

    if (selection && selection.rangeCount > 0) {
      range = selection.getRangeAt(0);
      // Verify the range is within our editor
      const container = range.commonAncestorContainer;
      if (!editor.contains(container.nodeType === Node.TEXT_NODE ? container.parentNode : container)) {
        // Selection is outside editor, create range at end
        range = document.createRange();
        range.selectNodeContents(editor);
        range.collapse(false);
      }
    } else {
      // No selection: create range at end of editor
      range = document.createRange();
      if (editor.childNodes.length > 0) {
        range.selectNodeContents(editor);
        range.collapse(false);
      } else {
        // Empty editor
        range.setStart(editor, 0);
        range.collapse(true);
      }
    }

    // Set selection to this range
    if (selection) {
      selection.removeAllRanges();
      selection.addRange(range);
    }

    // Insert image directly at the range
    try {
      range.deleteContents();
      range.insertNode(img);
      
      // Insert a line break after image for better editing
      const br = document.createElement("br");
      range.setStartAfter(img);
      range.collapse(true);
      range.insertNode(br);

      // Move cursor after the break
      if (selection) {
        range.setStartAfter(br);
        range.collapse(true);
        selection.removeAllRanges();
        selection.addRange(range);
      }
    } catch (err) {
      // Fallback: append to end
      console.warn("Direct insertion failed, appending to end:", err);
      editor.appendChild(img);
      const br = document.createElement("br");
      editor.appendChild(br);
    }

    // Keep editor focused
    editor.focus();
  };
  reader.onerror = (err) => {
    console.error("Failed to read image file:", err);
    alert("Failed to load image file.");
  };
  reader.readAsDataURL(file);
}

// ==== QUESTIONS ====
function validateQuestionInput(forUpdate = false) {
  const examInfo = validateExamInfo();
  if (!examInfo) return null;
  if (!ensureSetSelected()) return null;

  const setId = setSelectEl.value || state.currentSetId;
  if (!setId) {
    alert("Please select a question set.");
    return null;
  }

  const set = state.sets.find((s) => s.id === setId);
  if (!set) {
    alert("Selected set not found.");
    return null;
  }

  const path = qPathEl.value.trim();
  const answerType = Number(qAnswerTypeEl.value);
  const marks = Number(qMarksEl.value || 0);
  const needsContext = answerType === 1 && !!qNeedsContextEl.checked;

  const textHtml = getEditorHtml(qTextEditorEl);
  const textPlain = htmlToPlainText(textHtml);

  const markHtml = getEditorHtml(qMarkEditorEl);
  const markPlain = htmlToPlainText(markHtml);

  if (!path) {
    alert("Path is required.");
    return null;
  }
  if (!textPlain) {
    alert("Question text is required.");
    return null;
  }

  if (answerType === 1 || answerType === 2) {
    if (!markPlain) {
      alert("Mark scheme is required for answer types 1 and 2.");
      return null;
    }
    if (marks <= 0) {
      alert("Marks must be greater than 0 for answer types 1 and 2.");
      return null;
    }
  } else {
    // type 0: marks must be 0
    if (marks !== 0) {
      alert("Marks must be 0 for answer type 0 (context only).");
      return null;
    }
  }

  return {
    set,
    path,
    answerType,
    marks,
    needsContext,
    textHtml,
    markHtml,
  };
}

function addQuestion(resetEditorsAfter = true) {
  // Ensure we're not in edit mode
  if (state.editingQuestionId) {
    alert("Please cancel the current edit or update the question first.");
    return;
  }

  const validated = validateQuestionInput(false);
  if (!validated) return;
  const { set, path, answerType, marks, needsContext, textHtml, markHtml } =
    validated;

  const currentSetQuestions = state.questions.filter((q) => q.set_id === set.id);
  const nextOrder = currentSetQuestions.length + 1;

  const question = {
    uniqueid: uuid(),
    path,
    text_body: textHtml,
    answer_type: answerType,
    mark_scheme: markHtml,
    needs_context: needsContext,
    exam: state.exam.exam,
    subject: state.exam.subject,
    section: set.section,
    topic: "",
    order: nextOrder,
    marks,
    set_id: set.id,
    set_label: set.label,
  };

  state.questions.push(question);
  state.currentSetId = set.id;
  
  // Recalculate orders to ensure they're contiguous
  recalcOrders();
  
  refreshSetSelect();
  renderPreview();
  updateStatus("Question added.");

  if (resetEditorsAfter) {
    clearForm();
  }
}

function clearForm() {
  qPathEl.value = "";
  qTextEditorEl.innerHTML = "";
  qMarkEditorEl.innerHTML = "";
  qNeedsContextEl.checked = false;
  qMarksEl.value = "0";
  qPathEl.focus();
}

function loadQuestionIntoForm(question) {
  state.editingQuestionId = question.uniqueid;
  qFormTitleEl.textContent = "Edit question";

  setSelectEl.value = question.set_id;
  qPathEl.value = question.path;
  qAnswerTypeEl.value = String(question.answer_type);
  qNeedsContextEl.checked = !!question.needs_context;
  qMarksEl.value = String(question.marks ?? 0);
  qTextEditorEl.innerHTML = question.text_body || "";
  qMarkEditorEl.innerHTML = question.mark_scheme || "";

  updateNeedsContextVisibility();

  // Show Update/Cancel, hide Add buttons
  addBtn.hidden = true;
  addContinueBtn.hidden = true;
  updateBtn.hidden = false;
  cancelEditBtn.hidden = false;
}

function copyQuestionIntoForm(question) {
  // Copy mode: fill form but stay in Add mode
  state.editingQuestionId = null;
  qFormTitleEl.textContent = "Add question (copied)";

  setSelectEl.value = question.set_id;
  qPathEl.value = question.path;
  qAnswerTypeEl.value = String(question.answer_type);
  qNeedsContextEl.checked = !!question.needs_context;
  qMarksEl.value = String(question.marks ?? 0);
  qTextEditorEl.innerHTML = question.text_body || "";
  qMarkEditorEl.innerHTML = question.mark_scheme || "";

  updateNeedsContextVisibility();

  // Stay in Add mode
  addBtn.hidden = false;
  addContinueBtn.hidden = false;
  updateBtn.hidden = true;
  cancelEditBtn.hidden = true;
}

function clearEditMode() {
  state.editingQuestionId = null;
  qFormTitleEl.textContent = "Add question";
  addBtn.hidden = false;
  addContinueBtn.hidden = false;
  updateBtn.hidden = true;
  cancelEditBtn.hidden = true;
  clearForm();
}

function updateQuestion() {
  if (!state.editingQuestionId) {
    alert("No question is being edited.");
    return;
  }
  
  const question = state.questions.find((q) => q.uniqueid === state.editingQuestionId);
  if (!question) {
    alert("Question not found.");
    clearEditMode();
    return;
  }

  const validated = validateQuestionInput(true);
  if (!validated) return;
  const { set, path, answerType, marks, needsContext, textHtml, markHtml } =
    validated;

  // Preserve uniqueid - never change it
  const oldSetId = question.set_id;
  
  // Update all fields
  question.path = path;
  question.text_body = textHtml;
  question.answer_type = answerType;
  question.mark_scheme = markHtml;
  question.needs_context = needsContext;
  question.marks = marks;
  question.exam = state.exam.exam;
  question.subject = state.exam.subject;

  // If moved to a different set, update set info and recalc orders
  if (question.set_id !== set.id) {
    question.set_id = set.id;
    question.set_label = set.label;
    question.section = set.section;
    recalcOrders();
  } else {
    // Even if set didn't change, recalc to ensure order is correct
    recalcOrders();
  }

  renderPreview();
  updateStatus("Question updated.");
  clearEditMode();
}

function deleteQuestion(uniqueid) {
  const q = state.questions.find((x) => x.uniqueid === uniqueid);
  if (!q) return;
  const ok = window.confirm("Delete this question?");
  if (!ok) return;
  state.questions = state.questions.filter((x) => x.uniqueid !== uniqueid);
  recalcOrders();
  renderPreview();
  updateStatus("Question deleted.");
}

function moveQuestion(uniqueid, delta) {
  const question = state.questions.find((q) => q.uniqueid === uniqueid);
  if (!question) return;
  const setId = question.set_id;
  const setQuestions = state.questions
    .filter((q) => q.set_id === setId)
    .sort((a, b) => (a.order || 0) - (b.order || 0));

  const idx = setQuestions.findIndex((q) => q.uniqueid === uniqueid);
  const newIdx = idx + delta;
  if (newIdx < 0 || newIdx >= setQuestions.length) return;

  const [item] = setQuestions.splice(idx, 1);
  setQuestions.splice(newIdx, 0, item);

  // Write back order values
  setQuestions.forEach((q, i) => {
    q.order = i + 1;
  });

  renderPreview();
}

function recalcOrders() {
  const bySet = {};
  state.questions.forEach((q) => {
    if (!bySet[q.set_id]) bySet[q.set_id] = [];
    bySet[q.set_id].push(q);
  });
  Object.values(bySet).forEach((arr) => {
    arr
      .sort((a, b) => (a.order || 0) - (b.order || 0))
      .forEach((q, i) => {
        q.order = i + 1;
      });
  });
}

// ==== PREVIEW RENDERING WITH DRAG/DROP ====
let dragState = { questionId: null };

function renderPreview() {
  previewContainerEl.innerHTML = "";

  if (!state.sets.length || !state.questions.length) {
    previewEmptyEl.style.display = "block";
    exportExcelBtn.disabled = true;
    clearAllBtn.disabled = true;
    return;
  }

  previewEmptyEl.style.display = "none";
  exportExcelBtn.disabled = false;
  clearAllBtn.disabled = false;

  const bySet = {};
  state.questions.forEach((q) => {
    if (!bySet[q.set_id]) bySet[q.set_id] = [];
    bySet[q.set_id].push(q);
  });

  const setsSorted = [...state.sets].sort((a, b) =>
    a.label.localeCompare(b.label)
  );

  setsSorted.forEach((set) => {
    const questions = (bySet[set.id] || []).sort(
      (a, b) => (a.order || 0) - (b.order || 0)
    );
    if (!questions.length) return;

    const groupEl = document.createElement("div");
    groupEl.className = "set-group";
    groupEl.dataset.setId = set.id;

    const headerEl = document.createElement("div");
    headerEl.className = "set-group-header";

    const labelSpan = document.createElement("span");
    labelSpan.textContent = set.section
      ? `${set.label} (Section ${set.section})`
      : set.label;

    const toggleSpan = document.createElement("span");
    toggleSpan.className = "set-group-toggle";
    toggleSpan.textContent = "▼";

    headerEl.appendChild(labelSpan);
    headerEl.appendChild(toggleSpan);
    groupEl.appendChild(headerEl);

    const tableWrapper = document.createElement("div");
    tableWrapper.className = "questions-table-wrapper";

    const table = document.createElement("table");
    table.className = "questions-table-sm";
    const thead = document.createElement("thead");
    thead.innerHTML =
      "<tr><th>#</th><th>Path</th><th>Type</th><th>Question</th><th>Mark scheme</th><th>Marks</th><th></th></tr>";
    table.appendChild(thead);

    const tbody = document.createElement("tbody");
    questions.forEach((q) => {
      const tr = document.createElement("tr");
      tr.dataset.qid = q.uniqueid;
      tr.draggable = true;

      const orderTd = document.createElement("td");
      orderTd.textContent = String(q.order || "");

      const pathTd = document.createElement("td");
      pathTd.textContent = q.path;

      const typeTd = document.createElement("td");
      typeTd.textContent = formatAnswerTypeLabel(q.answer_type);

      const textTd = document.createElement("td");
      const textPreview = htmlToPlainText(q.text_body || "");
      textTd.textContent =
        textPreview.length > 80
          ? textPreview.slice(0, 80) + "…"
          : textPreview || "(empty)";

      const markTd = document.createElement("td");
      const markPreview = htmlToPlainText(q.mark_scheme || "");
      markTd.textContent =
        markPreview.length > 60
          ? markPreview.slice(0, 60) + "…"
          : markPreview || "(none)";

      const marksTd = document.createElement("td");
      marksTd.textContent = q.marks ?? "";

      const actionsTd = document.createElement("td");
      actionsTd.className = "actions-cell";

      const dragSpan = document.createElement("span");
      dragSpan.className = "drag-handle";
      dragSpan.textContent = "↕";

      const upBtn = document.createElement("button");
      upBtn.className = "btn tiny";
      upBtn.textContent = "↑";
      upBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        moveQuestion(q.uniqueid, -1);
      });

      const downBtn = document.createElement("button");
      downBtn.className = "btn tiny";
      downBtn.textContent = "↓";
      downBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        moveQuestion(q.uniqueid, 1);
      });

      const editBtn = document.createElement("button");
      editBtn.className = "btn tiny";
      editBtn.textContent = "Edit";
      editBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        loadQuestionIntoForm(q);
      });

      const copyBtn = document.createElement("button");
      copyBtn.className = "btn tiny";
      copyBtn.textContent = "Copy";
      copyBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        copyQuestionIntoForm(q);
      });

      const delBtn = document.createElement("button");
      delBtn.className = "btn tiny danger";
      delBtn.textContent = "Delete";
      delBtn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        deleteQuestion(q.uniqueid);
      });

      actionsTd.appendChild(dragSpan);
      actionsTd.appendChild(upBtn);
      actionsTd.appendChild(downBtn);
      actionsTd.appendChild(editBtn);
      actionsTd.appendChild(copyBtn);
      actionsTd.appendChild(delBtn);

      tr.appendChild(orderTd);
      tr.appendChild(pathTd);
      tr.appendChild(typeTd);
      tr.appendChild(textTd);
      tr.appendChild(markTd);
      tr.appendChild(marksTd);
      tr.appendChild(actionsTd);

      // Drag events
      tr.addEventListener("dragstart", (e) => {
        dragState.questionId = q.uniqueid;
        tr.classList.add("dragging");
        e.dataTransfer.effectAllowed = "move";
      });
      tr.addEventListener("dragend", () => {
        tr.classList.remove("dragging");
        dragState.questionId = null;
      });
      tr.addEventListener("dragover", (e) => {
        e.preventDefault();
      });
      tr.addEventListener("drop", (e) => {
        e.preventDefault();
        if (!dragState.questionId || dragState.questionId === q.uniqueid) return;
        handleDropOnQuestion(set.id, q.uniqueid);
      });

      tbody.appendChild(tr);
    });

    table.appendChild(tbody);
    tableWrapper.appendChild(table);
    groupEl.appendChild(tableWrapper);

    // Toggle collapse
    let expanded = true;
    headerEl.addEventListener("click", () => {
      expanded = !expanded;
      tableWrapper.style.display = expanded ? "block" : "none";
      toggleSpan.textContent = expanded ? "▼" : "►";
    });

    previewContainerEl.appendChild(groupEl);
  });
}

function handleDropOnQuestion(targetSetId, targetQuestionId) {
  const dragged = state.questions.find((q) => q.uniqueid === dragState.questionId);
  const target = state.questions.find((q) => q.uniqueid === targetQuestionId);
  if (!dragged || !target) return;

  // Move dragged question into target set and position just before target
  dragged.set_id = targetSetId;
  const setQuestions = state.questions.filter((q) => q.set_id === targetSetId);
  setQuestions.sort((a, b) => (a.order || 0) - (b.order || 0));

  const withoutDragged = setQuestions.filter((q) => q.uniqueid !== dragged.uniqueid);
  const idx = withoutDragged.findIndex((q) => q.uniqueid === target.uniqueid);
  if (idx === -1) return;

  withoutDragged.splice(idx, 0, dragged);
  withoutDragged.forEach((q, i) => {
    q.order = i + 1;
  });

  // Update set metadata
  const set = state.sets.find((s) => s.id === targetSetId);
  if (set) {
    dragged.set_label = set.label;
    dragged.section = set.section;
  }

  renderPreview();
}

// ==== EXCEL EXPORT WITH IMAGES AND FORMATTING ====
async function handleExportExcel() {
  const examInfo = validateExamInfo();
  if (!examInfo) return;
  if (!state.questions.length) {
    alert("No questions to export.");
    return;
  }

  if (typeof ExcelJS === "undefined") {
    alert("ExcelJS library not loaded. Please refresh the page.");
    return;
  }

  exportExcelBtn.disabled = true;
  updateStatus("Generating Excel file...");

  try {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Questions");

    // Recalculate orders
    recalcOrders();

    // Headers
    worksheet.columns = [
      { header: "uniqueid", key: "uniqueid", width: 15 },
      { header: "path", key: "path", width: 12 },
      { header: "text_body", key: "text_body", width: 50 },
      { header: "answer_type", key: "answer_type", width: 12 },
      { header: "mark_scheme", key: "mark_scheme", width: 50 },
      { header: "needs_context", key: "needs_context", width: 12 },
      { header: "exam", key: "exam", width: 10 },
      { header: "subject", key: "subject", width: 20 },
      { header: "set_label", key: "set_label", width: 15 },
      { header: "section", key: "section", width: 10 },
      { header: "topic", key: "topic", width: 12 },
      { header: "order", key: "order", width: 8 },
      { header: "marks", key: "marks", width: 8 },
    ];

    // Style header row
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE0E0E0" },
    };

    // Group by set and sort
    const bySet = {};
    state.questions.forEach((q) => {
      if (!bySet[q.set_id]) bySet[q.set_id] = [];
      bySet[q.set_id].push(q);
    });
    const setsSorted = [...state.sets].sort((a, b) =>
      a.label.localeCompare(b.label)
    );

    let rowNum = 2;
    const imagePromises = [];

    setsSorted.forEach((set) => {
      const questions = (bySet[set.id] || []).sort(
        (a, b) => (a.order || 0) - (b.order || 0)
      );

      questions.forEach((q) => {
        const row = worksheet.getRow(rowNum);

        // Basic fields
        row.getCell("uniqueid").value = q.uniqueid;
        row.getCell("path").value = q.path;
        row.getCell("answer_type").value = q.answer_type;
        row.getCell("needs_context").value = String(!!q.needs_context).toLowerCase();
        row.getCell("exam").value = q.exam || state.exam.exam;
        row.getCell("subject").value = q.subject || state.exam.subject;
        row.getCell("set_label").value = q.set_label || "";
        row.getCell("section").value = q.section || "";
        row.getCell("topic").value = q.topic || "";
        row.getCell("order").value = q.order ?? "";
        row.getCell("marks").value = q.marks ?? "";

        // Process text_body with formatting and images
        const textCell = row.getCell("text_body");
        processRichTextCell(textCell, q.text_body || "", imagePromises, rowNum, 3, worksheet);

        // Process mark_scheme with formatting and images
        const markCell = row.getCell("mark_scheme");
        processRichTextCell(markCell, q.mark_scheme || "", imagePromises, rowNum, 4, worksheet);

        rowNum++;
      });
    });

    // Wait for all images to be processed
    await Promise.all(imagePromises);

    // Generate file
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const examClean = (state.exam.exam || "questions").replace(/\s+/g, "_");
    const filename = `${examClean}_questions.xlsx`;

    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    updateStatus("Excel exported.");
  } catch (err) {
    console.error("Export error:", err);
    alert("Failed to export Excel file: " + err.message);
  } finally {
    exportExcelBtn.disabled = false;
  }
}

function processRichTextCell(cell, html, imagePromises, rowNum, colIndex, worksheet) {
  if (!html) {
    cell.value = "";
    cell.alignment = { wrapText: true, vertical: "top" };
    return;
  }

  // Parse HTML to extract text and formatting
  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = html;

  // Extract images first and store their positions
  const images = [];
  const imgElements = tempDiv.querySelectorAll("img");
  imgElements.forEach((img, idx) => {
    const src = img.getAttribute("src");
    if (src && src.startsWith("data:image")) {
      images.push({ src, index: idx });
      // Remove the <img> from the HTML so no placeholder text appears in the cell
      if (img.parentNode) {
        img.parentNode.removeChild(img);
      }
    }
  });

  // Build rich text array with formatting, preserving line breaks
  const richText = [];

  function processNode(node, inheritedFormat = {}) {
    if (node.nodeType === Node.TEXT_NODE) {
      const text = node.textContent;
      if (text) {
        // Preserve all text, including whitespace and newlines
        richText.push({
          text: text,
          font: { ...inheritedFormat },
        });
      }
    } else if (node.nodeType === Node.ELEMENT_NODE) {
      const tagName = node.tagName.toLowerCase();
      const newFormat = { ...inheritedFormat };

      if (tagName === "b" || tagName === "strong") {
        newFormat.bold = true;
      } else if (tagName === "i" || tagName === "em") {
        newFormat.italic = true;
      } else if (tagName === "br") {
        // Explicit line break
        richText.push({ text: "\n", font: inheritedFormat });
      } else if (tagName === "p") {
        // Paragraph: add newline before (except first)
        if (richText.length > 0) {
          richText.push({ text: "\n", font: inheritedFormat });
        }
      } else if (tagName === "div") {
        // Divs might contain line breaks
        if (richText.length > 0 && node.previousSibling) {
          richText.push({ text: "\n", font: inheritedFormat });
        }
      }

      // Check for monospace style
      const style = node.getAttribute("style") || "";
      if (style.includes("Courier") || style.includes("monospace") || 
          style.includes("font-family") && (style.includes("Courier") || style.includes("monospace"))) {
        newFormat.name = "Courier New";
      }

      // Process child nodes
      Array.from(node.childNodes).forEach((child) => {
        processNode(child, newFormat);
      });
    }
  }

  // Process all child nodes
  Array.from(tempDiv.childNodes).forEach((child) => {
    processNode(child);
  });

  // Merge consecutive text nodes with same formatting
  const merged = [];
  let current = null;
  
  richText.forEach((rt) => {
    const formatKey = JSON.stringify(rt.font || {});
    if (current && current.formatKey === formatKey) {
      current.text += rt.text;
    } else {
      if (current) {
        merged.push({ text: current.text, font: current.font });
      }
      current = { 
        text: rt.text, 
        font: rt.font || {}, 
        formatKey 
      };
    }
  });
  if (current) {
    merged.push({ text: current.text, font: current.font });
  }

  // Set cell value with rich text if we have formatting, otherwise plain text
  if (merged.length > 0) {
    // Check if we have any formatting
    const hasFormatting = merged.some(rt => 
      rt.font.bold || rt.font.italic || rt.font.name
    );
    
    if (hasFormatting) {
      cell.value = { richText: merged };
    } else {
      // No formatting, just use plain text with newlines
      const plainText = merged.map(rt => rt.text).join("");
      cell.value = plainText;
    }
  } else {
    // Fallback: extract plain text with line breaks
    let text = tempDiv.textContent || "";
    // Preserve newlines
    text = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n");
    cell.value = text;
  }

  // Add images to worksheet (positioned in/next to the cell)
  images.forEach((imgData, idx) => {
    const promise = addImageToWorksheet(worksheet, imgData.src, rowNum, colIndex, idx);
    imagePromises.push(promise);
  });

  // Enable word wrap and set alignment
  cell.alignment = { wrapText: true, vertical: "top" };
  
  // Set row height based on content and images
  // Keep rows without images compact, make image rows taller
  const baseHeight = 20;
  const imageHeight = images.length > 0 ? 150 : 0;
  cell.row.height = Math.max(cell.row.height || baseHeight, baseHeight + imageHeight);
}

async function addImageToWorksheet(worksheet, dataUrl, rowNum, colIndex, imageIndex) {
  try {
    // Convert data URL to buffer
    const base64Data = dataUrl.split(",")[1];
    if (!base64Data) {
      console.warn("No base64 data in image URL");
      return;
    }

    const binaryString = atob(base64Data);
    const imageBuffer = new Uint8Array(binaryString.length);
    for (let i = 0; i < binaryString.length; i++) {
      imageBuffer[i] = binaryString.charCodeAt(i);
    }

    // Get image dimensions
    const img = new Image();
    await new Promise((resolve, reject) => {
      img.onload = resolve;
      img.onerror = () => {
        console.warn("Failed to load image for dimensions");
        reject(new Error("Image load failed"));
      };
      img.src = dataUrl;
    });

    // Add image to workbook
    const imageId = worksheet.workbook.addImage({
      buffer: imageBuffer,
      extension: "png",
    });

    // Calculate position: place image in the cell area
    // ExcelJS uses 0-based column indices and row indices
    const col = colIndex - 1; // Convert to 0-based (colIndex is 1-based from processRichTextCell)
    const row = rowNum - 1; // Convert to 0-based (rowNum is 1-based)

    // Scale image to fit (max 200px height, maintain aspect ratio)
    const maxHeight = 200;
    const maxWidth = 300;
    let scale = Math.min(1, maxHeight / img.height, maxWidth / img.width);
    const width = Math.round(img.width * scale);
    const height = Math.round(img.height * scale);

    // Position image: place it in the cell, offset vertically for multiple images
    const yOffset = imageIndex * (height + 5); // Stack vertically with spacing

    worksheet.addImage(imageId, {
      tl: { col: col, row: row },
      ext: { width: width, height: height },
      editAs: "oneCell", // Anchor to one cell
    });
  } catch (err) {
    console.warn("Failed to add image to worksheet:", err);
    // Don't throw - continue with other images
  }
}

// ==== CLEAR ALL ====
function handleClearAll() {
  if (!state.sets.length && !state.questions.length) return;
  const ok = window.confirm(
    "This will remove all sets and questions currently entered. Proceed?"
  );
  if (!ok) return;
  state.sets = [];
  state.questions = [];
  state.currentSetId = null;
  state.editingQuestionId = null;
  renderSetsList();
  refreshSetSelect();
  renderPreview();
  clearEditMode();
  updateStatus("All entries cleared.");
}

// ==== STATUS ====
function updateStatus(message) {
  if (message) {
    statusEl.textContent = message;
    return;
  }
  const count = state.questions.length;
  const setCount = state.sets.length;
  statusEl.textContent = `${setCount} set(s) | ${count} question(s) entered`;
}

function updateNeedsContextVisibility() {
  const type = Number(qAnswerTypeEl.value);
  const isType1 = type === 1;
  if (isType1) {
    qNeedsContextEl.parentElement.style.display = "flex";
  } else {
    qNeedsContextEl.checked = false;
    qNeedsContextEl.parentElement.style.display = "none";
  }
}

// ==== INIT ====
function init() {
  // Rich-text toolbar wiring
  document.querySelectorAll(".editor-toolbar").forEach((toolbar) => {
    const targetId = toolbar.getAttribute("data-target");
    toolbar.querySelectorAll("[data-cmd]").forEach((btn) => {
      const cmd = btn.getAttribute("data-cmd");
      btn.addEventListener("click", (e) => {
        e.preventDefault();
        const editor = document.getElementById(targetId);
        applyEditorCommand(cmd, editor);
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

  // Drag and drop for editors
  [qTextEditorEl, qMarkEditorEl].forEach((editor) => {
    editor.addEventListener("dragover", (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
    editor.addEventListener("drop", (e) => {
      e.preventDefault();
      e.stopPropagation();
      const files = e.dataTransfer.files;
      if (files && files.length > 0) {
        const file = files[0];
        if (file.type.startsWith("image/")) {
          insertImageIntoEditor(editor.id, file);
        }
      }
    });
  });

  // Sets
  newSetBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleNewSet();
  });
  editSetBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleEditSet();
  });
  deleteSetBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleDeleteSet();
  });

  // Answer type change: toggle needs_context visibility
  qAnswerTypeEl.addEventListener("change", () => {
    updateNeedsContextVisibility();
  });
  updateNeedsContextVisibility();

  // Questions
  addBtn.addEventListener("click", (e) => {
    e.preventDefault();
    addQuestion(true);
  });
  addContinueBtn.addEventListener("click", (e) => {
    e.preventDefault();
    addQuestion(false);
  });
  updateBtn.addEventListener("click", (e) => {
    e.preventDefault();
    updateQuestion();
  });
  cancelEditBtn.addEventListener("click", (e) => {
    e.preventDefault();
    clearEditMode();
  });

  // Export & clear
  exportExcelBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleExportExcel();
  });
  clearAllBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleClearAll();
  });

  // Initialize button visibility (Add mode by default)
  clearEditMode();

  renderSetsList();
  refreshSetSelect();
  renderPreview();
  updateStatus();
}

document.addEventListener("DOMContentLoaded", init);
