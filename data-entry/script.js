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
  selectedQuestionId: null,
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
const exportCsvBtn = $("btn-export-csv");
const clearAllBtn = $("btn-clear-all");
const statusEl = $("status");
const moveUpGlobalBtn = $("btn-move-up-global");
const moveDownGlobalBtn = $("btn-move-down-global");

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

function quoteCsvField(value) {
  const s = String(value ?? "");
  const escaped = s.replace(/"/g, '""');
  return `"${escaped}"`;
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
  if (!/^[MN]\d{2}/.test(exam)) {
    alert("Exam code must start with M or N followed by two digits (e.g., M22, N23).");
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

function insertImageIntoEditor(editorId, file) {
  const editor = document.getElementById(editorId);
  if (!editor || !file) return;

  const reader = new FileReader();
  reader.onload = (e) => {
    const dataUrl = e.target.result;
    if (!dataUrl) return;

    const fileName = file.name || "attachment";
    const link = document.createElement("a");
    link.href = dataUrl;
    link.target = "_blank";
    link.textContent = fileName;

    const selection = window.getSelection();
    if (!selection || selection.rangeCount === 0) {
      editor.appendChild(link);
      editor.appendChild(document.createTextNode(" "));
      return;
    }
    const range = selection.getRangeAt(0);
    range.deleteContents();
    range.insertNode(link);
    range.collapse(false);
    range.insertNode(document.createTextNode(" "));
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
    section: set.section,
    topic: "",
    order: nextOrder,
    marks,
    set_id: set.id,
    set_label: set.label,
  };

  state.questions.push(question);
  state.currentSetId = set.id;
  refreshSetSelect();
  renderPreview();
  updateStatus("Question added.");

  if (resetEditorsAfter) {
    qPathEl.value = "";
    qTextEditorEl.innerHTML = "";
    qMarkEditorEl.innerHTML = "";
    qNeedsContextEl.checked = false;
    qMarksEl.value = "0";
    qPathEl.focus();
  }
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

  // Update needs_context visibility based on answer type
  updateNeedsContextVisibility();

  addBtn.hidden = true;
  addContinueBtn.hidden = true;
  updateBtn.hidden = false;
  cancelEditBtn.hidden = false;
}

function clearEditMode() {
  state.editingQuestionId = null;
  qFormTitleEl.textContent = "Add question";
  addBtn.hidden = false;
  addContinueBtn.hidden = false;
  updateBtn.hidden = true;
  cancelEditBtn.hidden = true;
  qTextEditorEl.innerHTML = "";
  qMarkEditorEl.innerHTML = "";
  qPathEl.value = "";
  qNeedsContextEl.checked = false;
  qMarksEl.value = "0";
}

function updateQuestion() {
  if (!state.editingQuestionId) return;
  const question = state.questions.find((q) => q.uniqueid === state.editingQuestionId);
  if (!question) {
    clearEditMode();
    return;
  }

  const validated = validateQuestionInput(true);
  if (!validated) return;
  const { set, path, answerType, marks, needsContext, textHtml, markHtml } =
    validated;

  // Preserve identifiers and metadata
  question.path = path;
  question.text_body = textHtml;
  question.answer_type = answerType;
  question.mark_scheme = markHtml;
  question.needs_context = needsContext;
  question.marks = marks;
  question.exam = state.exam.exam;

  // If moved to a different set, update set info and recalc orders
  if (question.set_id !== set.id) {
    question.set_id = set.id;
    question.set_label = set.label;
    question.section = set.section;
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

function moveSelectedQuestion(delta) {
  if (!state.selectedQuestionId) return;
  moveQuestion(state.selectedQuestionId, delta);
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
    exportCsvBtn.disabled = true;
    clearAllBtn.disabled = true;
    return;
  }

  previewEmptyEl.style.display = "none";
  exportCsvBtn.disabled = false;
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

      if (q.uniqueid === state.selectedQuestionId) {
        tr.classList.add("selected-row");
      }

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
        moveQuestion(q.uniqueid, -1);
      });

      const downBtn = document.createElement("button");
      downBtn.className = "btn tiny";
      downBtn.textContent = "↓";
      downBtn.addEventListener("click", (e) => {
        e.preventDefault();
        moveQuestion(q.uniqueid, 1);
      });

      const editBtn = document.createElement("button");
      editBtn.className = "btn tiny";
      editBtn.textContent = "Edit";
      editBtn.addEventListener("click", (e) => {
        e.preventDefault();
        loadQuestionIntoForm(q);
      });

      const delBtn = document.createElement("button");
      delBtn.className = "btn tiny danger";
      delBtn.textContent = "Delete";
      delBtn.addEventListener("click", (e) => {
        e.preventDefault();
        deleteQuestion(q.uniqueid);
      });

      actionsTd.appendChild(dragSpan);
      actionsTd.appendChild(upBtn);
      actionsTd.appendChild(downBtn);
      actionsTd.appendChild(editBtn);
      actionsTd.appendChild(delBtn);

      tr.appendChild(orderTd);
      tr.appendChild(pathTd);
      tr.appendChild(typeTd);
      tr.appendChild(textTd);
      tr.appendChild(markTd);
      tr.appendChild(marksTd);
      tr.appendChild(actionsTd);

      // Row selection for global move up/down
      tr.addEventListener("click", () => {
        state.selectedQuestionId = q.uniqueid;
        renderPreview();
        updateGlobalReorderButtons();
      });

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

// ==== CSV EXPORT ====
function handleExportCsv() {
  const examInfo = validateExamInfo();
  if (!examInfo) return;
  if (!state.questions.length) {
    alert("No questions to export.");
    return;
  }

  // Recalculate orders just in case
  recalcOrders();

  // Group by set and sort as per spec
  const bySet = {};
  state.questions.forEach((q) => {
    if (!bySet[q.set_id]) bySet[q.set_id] = [];
    bySet[q.set_id].push(q);
  });
  const setsSorted = [...state.sets].sort((a, b) =>
    a.label.localeCompare(b.label)
  );

  const header = [
    "uniqueid",
    "path",
    "text_body",
    "answer_type",
    "mark_scheme",
    "needs_context",
    "exam",
    "section",
    "topic",
    "order",
    "marks",
  ];

  const lines = [];
  lines.push(header.map(quoteCsvField).join(","));

  setsSorted.forEach((set) => {
    const questions = (bySet[set.id] || []).sort(
      (a, b) => (a.order || 0) - (b.order || 0)
    );
    questions.forEach((q) => {
      const row = [
        q.uniqueid,
        q.path,
        q.text_body || "",
        q.answer_type,
        q.mark_scheme || "",
        String(!!q.needs_context).toLowerCase(),
        q.exam || state.exam.exam,
        q.section || "",
        q.topic || "",
        q.order ?? "",
        q.marks ?? "",
      ];
      lines.push(row.map(quoteCsvField).join(","));
    });
  });

  const csv = lines.join("\r\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const examClean = (state.exam.exam || "questions").replace(/\s+/g, "_");
  const filename = `${examClean}_questions.csv`;

  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);

  updateStatus("CSV exported.");
  updateGlobalReorderButtons();
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
  renderSetsList();
  refreshSetSelect();
  renderPreview();
  updateStatus("All entries cleared.");
  state.selectedQuestionId = null;
  updateGlobalReorderButtons();
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

function updateGlobalReorderButtons() {
  const hasSelection = !!state.selectedQuestionId;
  moveUpGlobalBtn.disabled = !hasSelection;
  moveDownGlobalBtn.disabled = !hasSelection;
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
  exportCsvBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleExportCsv();
  });
  clearAllBtn.addEventListener("click", (e) => {
    e.preventDefault();
    handleClearAll();
  });

  // Global move up/down buttons
  moveUpGlobalBtn.addEventListener("click", (e) => {
    e.preventDefault();
    moveSelectedQuestion(-1);
  });
  moveDownGlobalBtn.addEventListener("click", (e) => {
    e.preventDefault();
    moveSelectedQuestion(1);
  });

  renderSetsList();
  refreshSetSelect();
  renderPreview();
  updateStatus();
  updateGlobalReorderButtons();
}

document.addEventListener("DOMContentLoaded", init);

