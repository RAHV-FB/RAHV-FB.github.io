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

// ==== LOCALSTORAGE PERSISTENCE ====
const STORAGE_KEY = "ib_question_entry_state";
const INSTRUCTIONS_STATE_KEY = "ib_question_entry_instructions_expanded";

function saveStateToStorage() {
  try {
    const dataToSave = {
      exam: state.exam,
      sets: state.sets,
      questions: state.questions,
      currentSetId: state.currentSetId,
      editingQuestionId: state.editingQuestionId,
    };
    const jsonString = JSON.stringify(dataToSave);
    localStorage.setItem(STORAGE_KEY, jsonString);
    console.log("State saved to localStorage:", {
      exam: dataToSave.exam,
      setsCount: dataToSave.sets.length,
      questionsCount: dataToSave.questions.length,
      currentSetId: dataToSave.currentSetId
    });
  } catch (err) {
    console.error("Failed to save state to localStorage:", err);
    // Check if it's a quota exceeded error
    if (err.name === 'QuotaExceededError' || err.code === 22) {
      alert("Storage quota exceeded. Please clear some data or use a different browser.");
    }
  }
}

function loadStateFromStorage() {
  try {
    const saved = localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      console.log("No saved state found in localStorage");
      return false;
    }

    const data = JSON.parse(saved);
    console.log("Loaded state from localStorage:", data);
    
    // Restore state with validation
    if (data.exam && typeof data.exam === 'object') {
      state.exam = { ...state.exam, ...data.exam };
    }
    if (Array.isArray(data.sets)) {
      state.sets = data.sets;
    }
    if (Array.isArray(data.questions)) {
      state.questions = data.questions;
    }
    if (data.currentSetId !== undefined && data.currentSetId !== null) {
      state.currentSetId = data.currentSetId;
    }
    if (data.editingQuestionId !== undefined) {
      state.editingQuestionId = data.editingQuestionId;
    }

    console.log("State restored:", state);
    return true;
  } catch (err) {
    console.error("Failed to load state from localStorage:", err);
    return false;
  }
}

function restoreFormFromState() {
  // Restore exam fields
  if (state.exam.subject) {
    // Check if the subject value exists in the select options
    const subjectOption = Array.from(subjectEl.options).find(
      opt => opt.textContent === state.exam.subject || opt.value === state.exam.subject
    );
    if (subjectOption) {
      subjectEl.value = subjectOption.value;
    } else {
      // If exact match not found, try to set by value directly
      subjectEl.value = state.exam.subject;
    }
  }
  if (state.exam.exam) {
    examEl.value = state.exam.exam;
  }
}

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
  saveStateToStorage();
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

  saveStateToStorage();
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

  saveStateToStorage();
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
  
  saveStateToStorage();
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

  saveStateToStorage();
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
  saveStateToStorage();
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

  saveStateToStorage();
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

  saveStateToStorage();
  renderPreview();
}

// ==== CSV EXPORT WITH HTML + IMAGE PLACEHOLDERS ====
async function handleExportExcel() {
  const examInfo = validateExamInfo();
  if (!examInfo) return;
  if (!state.questions.length) {
    alert("No questions to export.");
    return;
  }

  if (typeof JSZip === "undefined") {
    alert("JSZip library not loaded. Please refresh the page.");
    return;
  }

  exportExcelBtn.disabled = true;
  updateStatus("Generating Excel file...");

  try {
    const zip = new JSZip();

    // Recalculate orders
    recalcOrders();

    // CSV headers (no set_label column)
    const headers = [
      "uniqueid",
      "path",
      "text_body",
      "answer_type",
      "mark_scheme",
      "needs_context",
      "exam",
      "subject",
      "section",
      "topic",
      "order",
      "marks",
    ];

    // Group by set and sort
    const bySet = {};
    state.questions.forEach((q) => {
      if (!bySet[q.set_id]) bySet[q.set_id] = [];
      bySet[q.set_id].push(q);
    });
    const setsSorted = [...state.sets].sort((a, b) =>
      a.label.localeCompare(b.label)
    );

    const rows = [];
    rows.push(headers);

    setsSorted.forEach((set) => {
      const questions = (bySet[set.id] || []).sort(
        (a, b) => (a.order || 0) - (b.order || 0)
      );

      questions.forEach((q) => {
        // Transform HTML to keep formatting and replace each <img> with a placeholder
        // like //image:UUID..., while saving the decoded image to the zip.
        const textHtml = transformHtmlForCsv(q.text_body || "", zip);
        const markHtml = transformHtmlForCsv(q.mark_scheme || "", zip);

        const row = [
          q.uniqueid,
          q.path,
          textHtml,
          q.answer_type,
          markHtml,
          String(!!q.needs_context).toLowerCase(),
          q.exam || state.exam.exam,
          q.subject || state.exam.subject,
          q.section || "",
          q.topic || "",
          q.order ?? "",
          q.marks ?? "",
        ];
        rows.push(row);
      });
    });

    // Build CSV content (quote all fields, preserve HTML)
    const csvLines = rows.map((row) =>
      row
        .map((value) => {
          const str = value == null ? "" : String(value);
          const escaped = str.replace(/"/g, '""');
          return `"${escaped}"`;
        })
        .join(",")
    );
    const csvContent = csvLines.join("\r\n");

    const examClean = (state.exam.exam || "questions").replace(/\s+/g, "_");
    const csvFilename = `${examClean}_questions.csv`;

    // Put CSV inside zip
    zip.file(csvFilename, csvContent);

    // Generate zip (CSV + images folder)
    const zipBlob = await zip.generateAsync({ type: "blob" });
    const zipFilename = `${examClean}_questions_with_images.zip`;

    const url = URL.createObjectURL(zipBlob);
    const a = document.createElement("a");
    a.href = url;
    a.download = zipFilename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);

    updateStatus("CSV + images exported.");
  } catch (err) {
    console.error("Export error:", err);
    alert("Failed to export CSV file: " + err.message);
  } finally {
    exportExcelBtn.disabled = false;
  }
}

/**
 * Transform HTML for CSV export:
 * - Keep HTML tags so bold/italic/paragraphs are preserved.
 * - Replace each inline <img src="data:..."> with a unique placeholder string
 *   like //image:UUID...
 * - Decode each image and add it to the zip under images//image:UUID...
 */
function transformHtmlForCsv(html, zip) {
  if (!html) return "";

  const tempDiv = document.createElement("div");
  tempDiv.innerHTML = html;

  const imgElements = tempDiv.querySelectorAll("img");
  imgElements.forEach((img) => {
    const src = img.getAttribute("src") || "";
    if (!src.startsWith("data:image")) {
      // Non-data URLs are left as-is in the HTML
      return;
    }

    // Generate unique image ID placeholder
    const imageId = `//image:${uuid()}`;

    // Decode base64 data and add image file into the zip
    try {
      const base64Data = src.split(",")[1];
      if (!base64Data) {
        throw new Error("No base64 data in image src");
      }

      const binaryString = atob(base64Data);
      const len = binaryString.length;
      const bytes = new Uint8Array(len);
      for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
      }

      // Store under images/ with filename equal to the placeholder ID
      const imagePath = `images/${imageId}`;
      zip.file(imagePath, bytes);
    } catch (err) {
      console.warn("Failed to decode/export image for CSV:", err);
    }

    // Replace the <img> with the placeholder text node so the marker
    // appears in the correct position in the HTML flow.
    const placeholderNode = document.createTextNode(imageId);
    if (img.parentNode) {
      img.parentNode.replaceChild(placeholderNode, img);
    }
  });

  // Return HTML string with <img> tags replaced by //image:ID... markers
  return tempDiv.innerHTML;
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
  saveStateToStorage();
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

// ==== INSTRUCTIONS TOGGLE ====
function initInstructionsToggle() {
  const instructionsHeader = document.getElementById("instructions-header");
  const instructionsContent = document.getElementById("instructions-content");
  const instructionsToggle = document.getElementById("instructions-toggle");
  const instructionsToggleIcon = document.getElementById("instructions-toggle-icon");

  if (!instructionsHeader || !instructionsContent || !instructionsToggle) return;

  // Load saved state (default to expanded on first load)
  const savedExpanded = localStorage.getItem(INSTRUCTIONS_STATE_KEY);
  const isExpanded = savedExpanded === null ? true : savedExpanded === "true";

  // Set initial state
  instructionsContent.style.display = isExpanded ? "block" : "none";
  instructionsToggleIcon.textContent = isExpanded ? "▼" : "►";

  // Toggle on click
  instructionsHeader.addEventListener("click", () => {
    const currentlyExpanded = instructionsContent.style.display !== "none";
    const newExpanded = !currentlyExpanded;
    
    instructionsContent.style.display = newExpanded ? "block" : "none";
    instructionsToggleIcon.textContent = newExpanded ? "▼" : "►";
    
    // Save state
    localStorage.setItem(INSTRUCTIONS_STATE_KEY, String(newExpanded));
  });
}

// ==== INIT ====
function init() {
  // Initialize instructions toggle (must be before state load to avoid flicker)
  initInstructionsToggle();

  // Load saved state from localStorage
  const hasSavedState = loadStateFromStorage();
  if (hasSavedState) {
    // Validate currentSetId - ensure it matches an existing set
    if (state.currentSetId && !state.sets.find(s => s.id === state.currentSetId)) {
      console.warn("currentSetId doesn't match any set, resetting to first set or null");
      state.currentSetId = state.sets.length > 0 ? state.sets[0].id : null;
    }
    
    restoreFormFromState();
    // Restore UI after loading state
    renderSetsList();
    refreshSetSelect();
    renderPreview();
    updateStatus();
  }

  // Check localStorage availability
  if (typeof Storage === "undefined" || !window.localStorage) {
    console.error("localStorage is not available in this browser");
    alert("Warning: localStorage is not available. Your data will not be saved between sessions.");
  }

  // Auto-save exam fields when they change
  subjectEl.addEventListener("change", () => {
    state.exam.subject = subjectEl.value;
    saveStateToStorage();
  });
  // Use both input and change events for exam field to ensure it saves
  examEl.addEventListener("input", () => {
    state.exam.exam = examEl.value;
    saveStateToStorage();
  });
  examEl.addEventListener("change", () => {
    state.exam.exam = examEl.value;
    saveStateToStorage();
  });

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

  // Only render if state wasn't already loaded (to avoid double render)
  if (!hasSavedState) {
    renderSetsList();
    refreshSetSelect();
    renderPreview();
    updateStatus();
  }
}

document.addEventListener("DOMContentLoaded", init);
//redo
