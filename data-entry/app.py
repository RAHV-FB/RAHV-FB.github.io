"""
IB Question Entry System - Desktop GUI Application
A minimal desktop Python application for data entry teams to input and structure
exam-style questions from International Baccalaureate (IB) subjects.
"""

import base64
import csv
import io
import os
import re
import shutil
import sys
import uuid
from pathlib import Path
from typing import Dict, List, Optional
from urllib.parse import unquote

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QComboBox, QPushButton, QTableWidget, QTableWidgetItem,
    QDialog, QDialogButtonBox, QTextEdit, QRadioButton, QCheckBox, QGroupBox,
    QMessageBox, QFileDialog, QStatusBar, QHeaderView, QListWidget, QListWidgetItem,
    QToolBar, QTextBrowser, QTreeWidget, QTreeWidgetItem, QSpinBox, QScrollArea
)
from PyQt6.QtCore import Qt, QMimeData, pyqtSignal, QByteArray, QBuffer, QIODevice
from PyQt6.QtGui import QFont, QTextCharFormat, QTextCursor, QDragEnterEvent, QDropEvent, QImage, QPixmap

# IB Subjects list
SUBJECTS = [
    "Eng A: Lang and Lit",
    "Eng A: Literature",
    "Spanish B",
    "English B",
    "History",
    "Business Management",
    "Economics",
    "Psychology",
    "Biology",
    "Chemistry",
    "Physics",
    "ESS",
    "Math AA: Analysis and Approaches",
    "Math AI: Applications and Interpretations",
    "Computer Science",
    "Spanish A: Language and Literature",
    "Spanish A: Literature",
]

# Answer type definitions
ANSWER_TYPE_0 = "0 - No answer expected (context only)"
ANSWER_TYPE_1 = "1 - Open text answer"
ANSWER_TYPE_2 = "2 - Multiple choice"

ANSWER_TYPES = {
    0: ANSWER_TYPE_0,
    1: ANSWER_TYPE_1,
    2: ANSWER_TYPE_2,
}

# CSV Headers matching database schema
CSV_HEADERS = [
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
]


class RichTextEdit(QTextEdit):
    """Rich text editor with image support and formatting."""
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptRichText(True)
        self.setAcceptDrops(True)
    
    def insertFromMimeData(self, source):
        """Override to preserve line breaks when pasting."""
        if source.hasText():
            text = source.text()
            # Insert text preserving line breaks
            cursor = self.textCursor()
            cursor.insertText(text)
            self.setTextCursor(cursor)
        elif source.hasHtml():
            # If HTML is available, use it to preserve formatting
            html = source.html()
            cursor = self.textCursor()
            cursor.insertHtml(html)
            self.setTextCursor(cursor)
        else:
            super().insertFromMimeData(source)
        
    def dragEnterEvent(self, event: QDragEnterEvent):
        """Handle drag enter event for images."""
        if event.mimeData().hasUrls() or event.mimeData().hasImage():
            event.acceptProposedAction()
        else:
            super().dragEnterEvent(event)
    
    def dropEvent(self, event: QDropEvent):
        """Handle drop event to insert images."""
        if event.mimeData().hasImage():
            image = event.mimeData().imageData()
            self.insert_image(image)
            event.acceptProposedAction()
        elif event.mimeData().hasUrls():
            url = event.mimeData().urls()[0]
            if url.isLocalFile():
                file_path = url.toLocalFile()
                # Check if it's an image
                if file_path.lower().endswith(('.png', '.jpg', '.jpeg', '.gif', '.bmp', '.webp')):
                    pixmap = QPixmap(file_path)
                    if not pixmap.isNull():
                        self.insert_image(pixmap)
                        event.acceptProposedAction()
                        return
            super().dropEvent(event)
        else:
            super().dropEvent(event)
    
    def insert_image(self, image):
        """Insert image into the document as base64 data URI."""
        cursor = self.textCursor()
        
        # Convert QPixmap/QImage to QImage if needed
        if isinstance(image, QPixmap):
            image = image.toImage()
        
        # Convert to PNG bytes
        buffer = QByteArray()
        qbuffer = QBuffer(buffer)
        qbuffer.open(QIODevice.OpenModeFlag.WriteOnly)
        image.save(qbuffer, "PNG")
        
        # Convert to base64
        image_data = base64.b64encode(buffer.data()).decode('utf-8')
        data_uri = f"data:image/png;base64,{image_data}"
        
        # Insert HTML img tag
        html = f'<img src="{data_uri}" />'
        cursor.insertHtml(html)
        
        # Insert a space after the image so cursor can be positioned there
        cursor.insertText(" ")
        
        # Update the editor's cursor position and restore focus
        self.setTextCursor(cursor)
        self.setFocus()
        # Ensure the editor widget receives focus properly
        self.activateWindow()


class DragDropTreeWidget(QTreeWidget):
    """Tree widget with drag-and-drop reordering support."""
    
    items_reordered = pyqtSignal()
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setDragDropMode(QTreeWidget.DragDropMode.InternalMove)
        self.setDefaultDropAction(Qt.DropAction.MoveAction)
        self.setSelectionMode(QTreeWidget.SelectionMode.SingleSelection)
        self.setRootIsDecorated(True)
        self.setUniformRowHeights(True)
    
    def startDrag(self, supportedActions):
        """Override to ensure drag works properly."""
        item = self.currentItem()
        if item and item.parent() is not None:  # Only allow dragging question items, not groups
            super().startDrag(supportedActions)
    
    def dropEvent(self, event):
        """Override to emit signal after drop."""
        # Call parent drop event first
        super().dropEvent(event)
        # Emit signal after drop completes (with delay to ensure tree is updated)
        from PyQt6.QtCore import QTimer
        QTimer.singleShot(150, self.items_reordered.emit)


class QuestionSetDialog(QDialog):
    """Dialog for creating or editing a question set."""

    def __init__(self, parent=None, question_set: Optional[Dict] = None):
        super().__init__(parent)
        self.result = None
        self.question_set = question_set  # If provided, we're editing
        is_editing = question_set is not None
        
        self.setWindowTitle("Edit Question Set" if is_editing else "Create Question Set")
        self.setModal(True)
        self.resize(400, 150)

        layout = QVBoxLayout(self)

        # Set label
        label_layout = QHBoxLayout()
        label_layout.addWidget(QLabel("Set Label *:"))
        self.label_edit = QLineEdit()
        self.label_edit.setPlaceholderText("e.g., Question 10")
        if is_editing:
            self.label_edit.setText(question_set.get("label", ""))
        label_layout.addWidget(self.label_edit)
        layout.addLayout(label_layout)

        # Section
        section_layout = QHBoxLayout()
        section_layout.addWidget(QLabel("Section:"))
        self.section_combo = QComboBox()
        self.section_combo.addItems(["", "A", "B"])
        if is_editing:
            section = question_set.get("section", "")
            index = self.section_combo.findText(section)
            if index >= 0:
                self.section_combo.setCurrentIndex(index)
        section_layout.addWidget(self.section_combo)
        layout.addLayout(section_layout)

        # Buttons
        button_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        button_box.accepted.connect(self.save)
        button_box.rejected.connect(self.reject)
        layout.addWidget(button_box)

    def save(self):
        """Save the question set."""
        label = self.label_edit.text().strip()
        if not label:
            QMessageBox.critical(self, "Validation Error", "Set label is required.")
            return

        if self.question_set:
            # Editing existing set - preserve ID
            self.result = {
                "id": self.question_set["id"],
                "label": label,
                "section": self.section_combo.currentText().strip(),
            }
        else:
            # Creating new set
            self.result = {
                "id": str(uuid.uuid4()),
                "label": label,
                "section": self.section_combo.currentText().strip(),
            }
        self.accept()


class QuestionDialog(QDialog):
    """Dialog window for adding/editing questions (supports adding multiple)."""

    def __init__(self, parent=None, question_set_id: Optional[str] = None, question: Optional[Dict] = None):
        super().__init__(parent)
        self.result = None
        self.added_questions: List[Dict] = []
        self.question_set_id = question_set_id
        self.question = question
        self.is_adding_multiple = question is None and question_set_id is not None

        self.setWindowTitle("Add Question" if question is None else "Edit Question")
        self.setModal(True)
        self.resize(750, 800)

        self.setup_ui()

        # Hide mark scheme initially if answer type is 0
        self.on_answer_type_changed()

        # Load question data if editing
        if question:
            self.path_edit.setText(question.get("path", ""))
            # Load HTML if available, otherwise plain text
            text_content = question.get("text_body", "")
            if text_content.strip().startswith("<"):
                self.text_body_edit.setHtml(text_content)
            else:
                self.text_body_edit.setPlainText(text_content)
            
            answer_type = question.get("answer_type", 0)
            self.answer_type_buttons[answer_type].setChecked(True)
            
            # Load marks if available
            marks_value = question.get("marks", 0)
            self.marks_spinbox.setValue(marks_value)
            
            mark_content = question.get("mark_scheme", "")
            if mark_content.strip().startswith("<"):
                self.mark_scheme_edit.setHtml(mark_content)
            else:
                self.mark_scheme_edit.setPlainText(mark_content)
            
            # Update visibility based on answer type, then load needs_context if type 1
            self.on_answer_type_changed()
            
            # Load needs_context only for answer type 1 (on_answer_type_changed may have reset it)
            if answer_type == 1:
                self.needs_context_checkbox.setChecked(question.get("needs_context", False))
            else:
                self.needs_context_checkbox.setChecked(False)

    def setup_ui(self):
        """Setup the dialog UI components."""
        # Main layout - vertical box with scroll area and buttons
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(0, 0, 0, 0)
        
        # Create scroll area for the form content
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        
        # Create a container widget for all form fields
        form_widget = QWidget()
        layout = QVBoxLayout(form_widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Path field
        path_layout = QHBoxLayout()
        path_label = QLabel("Path *:")
        path_label.setMinimumWidth(120)
        self.path_edit = QLineEdit()
        self.path_edit.setPlaceholderText("e.g., 10 or 10.a.i")
        path_layout.addWidget(path_label)
        path_layout.addWidget(self.path_edit)
        layout.addLayout(path_layout)

        path_hint = QLabel("(e.g., 10 or 10.a.i)")
        path_hint.setStyleSheet("color: gray; font-size: 10px;")
        path_hint.setIndent(120)
        layout.addWidget(path_hint)

        # Answer type
        answer_type_label = QLabel("Answer Type *:")
        answer_type_label.setMinimumWidth(120)
        layout.addWidget(answer_type_label)

        self.answer_type_buttons = {}
        answer_type_group = QGroupBox()
        answer_type_layout = QVBoxLayout(answer_type_group)
        
        for value, label in ANSWER_TYPES.items():
            radio = QRadioButton(label)
            if value == 0:
                radio.setChecked(True)
            self.answer_type_buttons[value] = radio
            radio.toggled.connect(self.on_answer_type_changed)
            answer_type_layout.addWidget(radio)
        
        layout.addWidget(answer_type_group)

        # Needs context checkbox (only for answer type 1)
        self.needs_context_checkbox = QCheckBox("Needs context")
        layout.addWidget(self.needs_context_checkbox)
        # Initially hidden - will be shown only for answer type 1
        self.needs_context_checkbox.setVisible(False)

        # Text body with formatting toolbar
        text_label = QLabel("Question Text *:")
        layout.addWidget(text_label)
        
        # Formatting toolbar for text body
        text_toolbar = QToolBar()
        bold_btn = QPushButton("Bold")
        bold_btn.setCheckable(True)
        bold_btn.clicked.connect(lambda: self.toggle_bold(self.text_body_edit))
        text_toolbar.addWidget(bold_btn)
        
        courier_btn = QPushButton("Courier")
        courier_btn.clicked.connect(lambda: self.apply_courier(self.text_body_edit))
        text_toolbar.addWidget(courier_btn)
        
        normal_btn = QPushButton("Normal")
        normal_btn.clicked.connect(lambda: self.apply_normal(self.text_body_edit))
        text_toolbar.addWidget(normal_btn)
        
        insert_image_btn = QPushButton("Insert Image")
        insert_image_btn.clicked.connect(lambda: self.insert_image(self.text_body_edit))
        text_toolbar.addWidget(insert_image_btn)
        
        layout.addWidget(text_toolbar)
        
        self.text_body_edit = RichTextEdit()
        self.text_body_edit.setMinimumHeight(150)
        layout.addWidget(self.text_body_edit)

        # Marks field (conditional - initially hidden)
        self.marks_label = QLabel("Marks *:")
        self.marks_spinbox = QSpinBox()
        self.marks_spinbox.setMinimum(0)
        self.marks_spinbox.setMaximum(100)
        self.marks_spinbox.setValue(0)
        marks_layout = QHBoxLayout()
        marks_layout.addWidget(self.marks_label)
        marks_layout.addWidget(self.marks_spinbox)
        marks_layout.addStretch()
        self.marks_widget = QWidget()
        self.marks_widget.setLayout(marks_layout)
        layout.addWidget(self.marks_widget)
        # Initially hidden (will be shown when answer type is 1 or 2)
        self.marks_widget.setVisible(False)

        # Mark scheme (conditional - initially hidden) with formatting toolbar
        self.mark_scheme_label = QLabel("Mark Scheme *:")
        self.mark_scheme_toolbar = QToolBar()
        mark_bold_btn = QPushButton("Bold")
        mark_bold_btn.setCheckable(True)
        mark_bold_btn.clicked.connect(lambda: self.toggle_bold(self.mark_scheme_edit))
        self.mark_scheme_toolbar.addWidget(mark_bold_btn)
        
        mark_courier_btn = QPushButton("Courier")
        mark_courier_btn.clicked.connect(lambda: self.apply_courier(self.mark_scheme_edit))
        self.mark_scheme_toolbar.addWidget(mark_courier_btn)
        
        mark_normal_btn = QPushButton("Normal")
        mark_normal_btn.clicked.connect(lambda: self.apply_normal(self.mark_scheme_edit))
        self.mark_scheme_toolbar.addWidget(mark_normal_btn)
        
        mark_insert_image_btn = QPushButton("Insert Image")
        mark_insert_image_btn.clicked.connect(lambda: self.insert_image(self.mark_scheme_edit))
        self.mark_scheme_toolbar.addWidget(mark_insert_image_btn)
        
        self.mark_scheme_edit = RichTextEdit()
        self.mark_scheme_edit.setMinimumHeight(120)
        layout.addWidget(self.mark_scheme_label)
        layout.addWidget(self.mark_scheme_toolbar)
        layout.addWidget(self.mark_scheme_edit)
        # Initially hidden (will be shown when answer type is 1 or 2)
        self.mark_scheme_label.setVisible(False)
        self.mark_scheme_toolbar.setVisible(False)
        self.mark_scheme_edit.setVisible(False)

        # Add stretch at the end to push content up
        layout.addStretch()
        
        # Set the form widget as the scroll area's widget
        scroll_area.setWidget(form_widget)
        main_layout.addWidget(scroll_area)

        # Buttons - outside scroll area so they're always visible
        button_layout = QHBoxLayout()
        button_layout.setContentsMargins(10, 5, 10, 10)
        
        if self.is_adding_multiple:
            add_another_btn = QPushButton("Add and Continue")
            add_another_btn.clicked.connect(self.add_and_continue)
            button_layout.addWidget(add_another_btn)

        button_box = QDialogButtonBox()
        if self.question is None:
            button_box.addButton("Add", QDialogButtonBox.ButtonRole.AcceptRole)
        else:
            button_box.addButton("Update", QDialogButtonBox.ButtonRole.AcceptRole)
        button_box.addButton("Cancel", QDialogButtonBox.ButtonRole.RejectRole)
        button_box.accepted.connect(self.save)
        button_box.rejected.connect(self.reject)
        button_layout.addWidget(button_box)
        
        main_layout.addLayout(button_layout)

    def toggle_bold(self, editor: QTextEdit):
        """Toggle bold formatting at cursor position."""
        cursor = editor.textCursor()
        fmt = QTextCharFormat()
        if cursor.charFormat().fontWeight() == QFont.Weight.Bold:
            fmt.setFontWeight(QFont.Weight.Normal)
        else:
            fmt.setFontWeight(QFont.Weight.Bold)
        cursor.mergeCharFormat(fmt)
        editor.setTextCursor(cursor)
    
    def apply_courier(self, editor: RichTextEdit):
        """Apply Courier font to selected text, preserving font size."""
        cursor = editor.textCursor()
        
        # If no selection, select word at cursor
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        
        # Get current format to preserve font size
        current_format = cursor.charFormat()
        current_font_size = current_format.fontPointSize()
        
        # Apply Courier font while preserving size
        fmt = QTextCharFormat()
        fmt.setFontFamily("Courier")
        # Preserve the font size - if no size is set, use editor's default
        if current_font_size > 0:
            fmt.setFontPointSize(current_font_size)
        else:
            # Use the editor's default font size
            default_font = editor.font()
            fmt.setFontPointSize(default_font.pointSizeF() if default_font.pointSizeF() > 0 else 12)
        
        cursor.mergeCharFormat(fmt)
        editor.setTextCursor(cursor)
    
    def apply_normal(self, editor: RichTextEdit):
        """Apply normal/default font to selected text, preserving font size."""
        cursor = editor.textCursor()
        
        # If no selection, select word at cursor
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        
        # Get current format to preserve font size and other properties
        current_format = cursor.charFormat()
        current_font_size = current_format.fontPointSize()
        
        # Get editor's default font
        default_font = editor.font()
        default_font_family = default_font.family()
        default_font_size = default_font.pointSizeF() if default_font.pointSizeF() > 0 else current_format.fontPointSize()
        
        # Use current size if available, otherwise use editor's default size
        font_size_to_use = current_font_size if current_font_size > 0 else default_font_size
        if font_size_to_use <= 0:
            font_size_to_use = 12  # Fallback to 12pt if nothing else works
        
        # Reset to default font while preserving size
        fmt = QTextCharFormat()
        fmt.setFontFamily(default_font_family)
        fmt.setFontPointSize(font_size_to_use)
        
        cursor.mergeCharFormat(fmt)
        editor.setTextCursor(cursor)

    def insert_image(self, editor: RichTextEdit):
        """Open file dialog to insert an image."""
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Select Image",
            "",
            "Image Files (*.png *.jpg *.jpeg *.gif *.bmp *.webp);;All Files (*)"
        )
        if filename:
            pixmap = QPixmap(filename)
            if not pixmap.isNull():
                editor.insert_image(pixmap)

    def on_answer_type_changed(self):
        """Show/hide mark scheme, marks, and needs_context fields based on answer type."""
        answer_type = self.get_answer_type()
        if answer_type in (1, 2):
            self.marks_widget.setVisible(True)
            self.mark_scheme_label.setVisible(True)
            self.mark_scheme_toolbar.setVisible(True)
            self.mark_scheme_edit.setVisible(True)
        else:
            self.marks_widget.setVisible(False)
            self.mark_scheme_label.setVisible(False)
            self.mark_scheme_toolbar.setVisible(False)
            self.mark_scheme_edit.setVisible(False)
        
        # Show needs_context checkbox only for answer type 1
        if answer_type == 1:
            self.needs_context_checkbox.setVisible(True)
            # Leave unchecked by default (optional field)
            if not self.question:  # Only reset if not editing
                self.needs_context_checkbox.setChecked(False)
        else:
            self.needs_context_checkbox.setVisible(False)
            # Set to False for types 0 and 2
            self.needs_context_checkbox.setChecked(False)

    def get_answer_type(self):
        """Get the selected answer type."""
        for value, radio in self.answer_type_buttons.items():
            if radio.isChecked():
                return value
        return 0

    def validate_and_get_question(self):
        """Validate and return question data."""
        path = self.path_edit.text().strip()
        # Get HTML content (preserves formatting and images)
        text_body = self.text_body_edit.toHtml().strip()
        answer_type = self.get_answer_type()
        mark_scheme = self.mark_scheme_edit.toHtml().strip() if answer_type in (1, 2) else ""

        # Validation
        if not path:
            QMessageBox.critical(self, "Validation Error", "Path is required.")
            return None
        
        # Check if text body has actual content (not just empty HTML tags)
        plain_text = self.text_body_edit.toPlainText().strip()
        if not plain_text:
            QMessageBox.critical(self, "Validation Error", "Question text is required.")
            return None
        
        if answer_type in (1, 2):
            mark_plain_text = self.mark_scheme_edit.toPlainText().strip()
            if not mark_plain_text:
                QMessageBox.critical(
                    self, "Validation Error", f"Mark scheme is required for answer type {answer_type}."
                )
                return None
            
            # Validate marks
            marks_value = self.marks_spinbox.value()
            if marks_value <= 0:
                QMessageBox.critical(
                    self, "Validation Error", "Marks must be greater than 0 for answer types 1 and 2."
                )
                return None
        
        # Determine needs_context value
        # For type 1: use checkbox value (can be True or False - optional)
        # For types 0 and 2: always False
        needs_context = self.needs_context_checkbox.isChecked() if answer_type == 1 else False

        question_data = {
            "path": path,
            "text_body": text_body,
            "answer_type": answer_type,
            "mark_scheme": mark_scheme,
            "needs_context": needs_context,
            "marks": self.marks_spinbox.value() if answer_type in (1, 2) else 0,
        }

        # Preserve uniqueid and order if editing
        if self.question:
            question_data["uniqueid"] = self.question.get("uniqueid")
            question_data["order"] = self.question.get("order")

        return question_data

    def save(self):
        """Save the question and close dialog."""
        question_data = self.validate_and_get_question()
        if question_data:
            self.result = question_data
            self.accept()

    def add_and_continue(self):
        """Add current question and clear form for next question."""
        question_data = self.validate_and_get_question()
        if question_data:
            self.added_questions.append(question_data)
            # Clear form for next question
            self.path_edit.clear()
            self.text_body_edit.clear()
            self.mark_scheme_edit.clear()
            # Reset to answer type 0 (this will hide needs_context checkbox and reset it)
            self.answer_type_buttons[0].setChecked(True)
            self.on_answer_type_changed()  # This will handle hiding and resetting needs_context
            # Focus on path field
            self.path_edit.setFocus()
            
            QMessageBox.information(self, "Added", f"Question added. You can add another question or click 'Add' to finish.")


class IBQuestionEntryApp(QMainWindow):
    """Main application window for IB Question Entry System."""

    def __init__(self):
        super().__init__()
        self.questions: List[Dict] = []
        self.question_sets: List[Dict] = []
        self.current_set_id: Optional[str] = None
        self.init_ui()

    def init_ui(self):
        """Initialize the main application UI."""
        self.setWindowTitle("IB Question Entry System")
        self.setGeometry(100, 100, 1200, 750)
        self.setMinimumSize(900, 600)

        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # Top frame for exam metadata
        exam_group = QGroupBox("Exam Information")
        exam_layout = QHBoxLayout()

        exam_layout.addWidget(QLabel("Subject *:"))
        self.subject_combo = QComboBox()
        self.subject_combo.addItems(SUBJECTS)
        self.subject_combo.setCurrentText("Computer Science")
        self.subject_combo.setMinimumWidth(250)
        exam_layout.addWidget(self.subject_combo)

        exam_layout.addWidget(QLabel("Exam *:"))
        self.exam_edit = QLineEdit()
        self.exam_edit.setMinimumWidth(200)
        exam_layout.addWidget(self.exam_edit)

        exam_layout.addStretch()
        exam_group.setLayout(exam_layout)
        layout.addWidget(exam_group)

        # Question Sets section
        sets_group = QGroupBox("Question Sets")
        sets_layout = QHBoxLayout()

        # Sets list
        sets_list_label = QLabel("Question Sets:")
        sets_list_label.setMinimumWidth(100)
        sets_layout.addWidget(sets_list_label)

        self.sets_list = QListWidget()
        self.sets_list.setMaximumWidth(250)
        self.sets_list.itemSelectionChanged.connect(self.on_set_selected)
        sets_layout.addWidget(self.sets_list)

        # Set management buttons
        sets_buttons_layout = QVBoxLayout()
        create_set_btn = QPushButton("Create New Set")
        create_set_btn.clicked.connect(self.create_question_set)
        sets_buttons_layout.addWidget(create_set_btn)

        edit_set_btn = QPushButton("Edit Selected Set")
        edit_set_btn.clicked.connect(self.edit_question_set)
        sets_buttons_layout.addWidget(edit_set_btn)

        delete_set_btn = QPushButton("Delete Selected Set")
        delete_set_btn.clicked.connect(self.delete_question_set)
        sets_buttons_layout.addWidget(delete_set_btn)
        sets_buttons_layout.addStretch()

        sets_layout.addLayout(sets_buttons_layout)
        sets_layout.addStretch()
        sets_group.setLayout(sets_layout)
        layout.addWidget(sets_group)

        # Button frame
        button_layout = QHBoxLayout()
        add_btn = QPushButton("Add Question(s)")
        add_btn.clicked.connect(self.add_question)
        button_layout.addWidget(add_btn)

        edit_btn = QPushButton("Edit Selected")
        edit_btn.clicked.connect(self.edit_question)
        button_layout.addWidget(edit_btn)

        delete_btn = QPushButton("Delete Selected")
        delete_btn.clicked.connect(self.delete_question)
        button_layout.addWidget(delete_btn)

        button_layout.addWidget(QLabel("  "))  # Spacer
        
        move_up_btn = QPushButton("Move Up")
        move_up_btn.clicked.connect(self.move_question_up)
        button_layout.addWidget(move_up_btn)

        move_down_btn = QPushButton("Move Down")
        move_down_btn.clicked.connect(self.move_question_down)
        button_layout.addWidget(move_down_btn)

        export_btn = QPushButton("Export CSV")
        export_btn.clicked.connect(self.export_csv)
        button_layout.addWidget(export_btn)

        button_layout.addStretch()
        layout.addLayout(button_layout)

        # Preview tree with collapsible groups and drag-and-drop support
        tree_group = QGroupBox("Question Preview")
        tree_layout = QVBoxLayout()

        self.tree = DragDropTreeWidget()
        self.tree.setHeaderLabels(["Order", "Path", "Answer Type", "Question Text Preview", "Mark Scheme Preview"])
        self.tree.setAlternatingRowColors(True)
        
        # Configure header
        header = self.tree.header()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        
        # Connect reorder signal
        self.tree.items_reordered.connect(self.rebuild_questions_from_tree)
        
        tree_layout.addWidget(self.tree)
        tree_group.setLayout(tree_layout)
        layout.addWidget(tree_group)

        # Status bar
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.update_status()

    def rebuild_questions_from_tree(self):
        """Rebuild questions list from tree order after drag-and-drop - preserves all data."""
        # Create a mapping of uniqueid -> full question data to preserve everything
        questions_dict = {q["uniqueid"]: q.copy() for q in self.questions}
        
        # Map set labels to set_ids
        set_label_to_id = {qset["label"]: qset["id"] for qset in self.question_sets}
        
        # Rebuild questions list grouped by set, maintaining order within each set
        questions_by_set = {}
        seen_uniqueids = set()
        
        # Collect questions from tree, organized by set
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            set_item = root.child(i)
            set_label_text = set_item.text(0)  # Get the group label
            
            # Extract set_id from the label (it should match one of our sets)
            set_id = None
            for label, sid in set_label_to_id.items():
                if label in set_label_text or set_label_text.startswith(label):
                    set_id = sid
                    break
            
            if set_id is None:
                continue
            
            if set_id not in questions_by_set:
                questions_by_set[set_id] = []
            
            # Collect questions from this set in tree order
            for j in range(set_item.childCount()):
                question_item = set_item.child(j)
                uniqueid = question_item.data(0, Qt.ItemDataRole.UserRole) if question_item else None
                if uniqueid and uniqueid in questions_dict and uniqueid not in seen_uniqueids:
                    question = questions_dict[uniqueid].copy()
                    # Update set_id if question was moved to a different set
                    question["set_id"] = set_id
                    qset = next((s for s in self.question_sets if s["id"] == set_id), None)
                    if qset:
                        question["set_label"] = qset["label"]
                        question["section"] = qset["section"]
                    questions_by_set[set_id].append(question)
                    seen_uniqueids.add(uniqueid)
        
        # Add any questions that weren't in the tree (safety check)
        for uniqueid, question in questions_dict.items():
            if uniqueid not in seen_uniqueids:
                set_id = question.get("set_id")
                if set_id not in questions_by_set:
                    questions_by_set[set_id] = []
                questions_by_set[set_id].append(question.copy())
        
        # Rebuild questions list maintaining set grouping
        new_questions_order = []
        for qset in self.question_sets:
            set_id = qset["id"]
            if set_id in questions_by_set:
                new_questions_order.extend(questions_by_set[set_id])
        
        # Update questions list
        self.questions = new_questions_order
        # Update all order numbers based on positions within each set
        self.update_order_numbers()
        # Update preview
        self.update_preview()

    def update_order_numbers(self):
        """Update order numbers for all questions based on their position within each set."""
        # Group questions by set_id
        questions_by_set = {}
        for question in self.questions:
            set_id = question.get("set_id")
            if set_id not in questions_by_set:
                questions_by_set[set_id] = []
            questions_by_set[set_id].append(question)
        
        # Update order numbers within each set (starting from 1 for each set)
        for set_id, set_questions in questions_by_set.items():
            # Sort by current order to preserve relative ordering
            sorted_questions = sorted(set_questions, key=lambda x: x.get("order", 999999))
            for idx, question in enumerate(sorted_questions, start=1):
                question["order"] = idx

    def create_question_set(self):
        """Create a new question set."""
        dialog = QuestionSetDialog(self)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            self.question_sets.append(dialog.result)
            self.update_sets_list()
            # Select the newly created set
            self.current_set_id = dialog.result["id"]
            self.select_set_in_list(self.current_set_id)

    def edit_question_set(self):
        """Edit the selected question set."""
        selected_items = self.sets_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question set to edit.")
            return
        
        item = selected_items[0]
        set_id = item.data(Qt.ItemDataRole.UserRole)
        question_set = next((s for s in self.question_sets if s["id"] == set_id), None)
        
        if not question_set:
            QMessageBox.warning(self, "Error", "Question set not found.")
            return
        
        dialog = QuestionSetDialog(self, question_set=question_set)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            # Update the question set in the list
            for i, qset in enumerate(self.question_sets):
                if qset["id"] == set_id:
                    self.question_sets[i] = dialog.result
                    # Update section for all questions in this set
                    for question in self.questions:
                        if question.get("set_id") == set_id:
                            question["section"] = dialog.result["section"]
                    break
            
            self.update_sets_list()
            self.update_preview()
            self.select_set_in_list(set_id)

    def delete_question_set(self):
        """Delete the selected question set."""
        selected_items = self.sets_list.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question set to delete.")
            return

        item = selected_items[0]
        set_id = item.data(Qt.ItemDataRole.UserRole)

        reply = QMessageBox.question(
            self, "Confirm Delete",
            "Are you sure you want to delete this question set and all its questions?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            # Remove the set
            self.question_sets = [s for s in self.question_sets if s["id"] != set_id]
            # Remove all questions from this set
            self.questions = [q for q in self.questions if q.get("set_id") != set_id]
            self.current_set_id = None
            self.update_sets_list()
            self.update_order_numbers()
            self.update_preview()
            self.update_status()

    def update_sets_list(self):
        """Update the question sets list widget."""
        self.sets_list.clear()
        for qset in self.question_sets:
            section_text = f" (Section {qset['section']})" if qset['section'] else ""
            item_text = f"{qset['label']}{section_text}"
            item = QListWidgetItem(item_text)
            item.setData(Qt.ItemDataRole.UserRole, qset["id"])
            self.sets_list.addItem(item)

    def select_set_in_list(self, set_id: str):
        """Select a set in the list by ID."""
        for i in range(self.sets_list.count()):
            item = self.sets_list.item(i)
            if item.data(Qt.ItemDataRole.UserRole) == set_id:
                self.sets_list.setCurrentItem(item)
                break

    def on_set_selected(self):
        """Handle question set selection."""
        selected_items = self.sets_list.selectedItems()
        if selected_items:
            self.current_set_id = selected_items[0].data(Qt.ItemDataRole.UserRole)
        else:
            self.current_set_id = None

    def add_question(self):
        """Open dialog to add question(s)."""
        exam_name = self.exam_edit.text().strip()
        if not exam_name:
            QMessageBox.critical(self, "Error", "Please enter an exam name first.")
            return
        
        # Validate exam name format (M or N followed by two digits)
        if not re.match(r'^[MN]\d{2}', exam_name):
            reply = QMessageBox.question(
                self,
                "Exam Name Warning",
                "Exam name should start with M or N followed by two digits (e.g., M22, N23).\n"
                "Do you want to continue anyway?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.No:
                return

        if not self.current_set_id:
            QMessageBox.warning(self, "No Set Selected", "Please select or create a question set first.")
            return

        dialog = QuestionDialog(self, question_set_id=self.current_set_id)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            # Add any questions that were added via "Add and Continue"
            added_count = len(dialog.added_questions)
            
            # Add the final question if it exists
            if dialog.result:
                dialog.added_questions.append(dialog.result)
                added_count += 1

            if dialog.added_questions:
                # Get the question set for section info
                qset = next((s for s in self.question_sets if s["id"] == self.current_set_id), None)
                if not qset:
                    return

                # Get current questions in this set to determine next order number
                current_set_questions = [q for q in self.questions if q.get("set_id") == self.current_set_id]
                next_order_in_set = len(current_set_questions) + 1
                
                for question_data in dialog.added_questions:
                    question = question_data.copy()
                    question["uniqueid"] = str(uuid.uuid4())
                    question["exam"] = self.exam_edit.text().strip()
                    question["section"] = qset["section"]
                    question["topic"] = ""  # Leave topic empty/null
                    question["set_id"] = self.current_set_id
                    question["set_label"] = qset["label"]
                    # Assign order number within this set
                    question["order"] = next_order_in_set
                    next_order_in_set += 1
                    self.questions.append(question)

                self.update_preview()
                self.update_status()
                QMessageBox.information(self, "Success", f"{added_count} question(s) added successfully.")

    def edit_question(self):
        """Edit the selected question."""
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question to edit.")
            return

        item = selected_items[0]
        # Make sure it's a question item, not a group item
        if item.parent() is None or item.parent().parent() is not None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a question item (not a group).")
            return

        uniqueid = item.data(0, Qt.ItemDataRole.UserRole)
        if not uniqueid:
            return

        # Find the question
        question = next((q for q in self.questions if q["uniqueid"] == uniqueid), None)
        if not question:
            QMessageBox.critical(self, "Error", "Question not found.")
            return

        dialog = QuestionDialog(self, question=question)
        if dialog.exec() == QDialog.DialogCode.Accepted and dialog.result:
            # Update question while preserving uniqueid, metadata, and order
            updated = dialog.result
            question.update({
                "path": updated["path"],
                "text_body": updated["text_body"],
                "answer_type": updated["answer_type"],
                "mark_scheme": updated["mark_scheme"],
                "needs_context": updated["needs_context"],
            })
            self.update_preview()
            self.update_status()
            QMessageBox.information(self, "Success", "Question updated successfully.")

    def delete_question(self):
        """Delete the selected question."""
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question to delete.")
            return

        item = selected_items[0]
        # Make sure it's a question item, not a group item
        if item.parent() is None or item.parent().parent() is not None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a question item (not a group).")
            return

        uniqueid = item.data(0, Qt.ItemDataRole.UserRole)
        if not uniqueid:
            return

        reply = QMessageBox.question(
            self, "Confirm Delete", "Are you sure you want to delete this question?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.questions = [q for q in self.questions if q["uniqueid"] != uniqueid]
            self.update_order_numbers()
            self.update_preview()
            self.update_status()
            QMessageBox.information(self, "Success", "Question deleted successfully.")

    def move_question_up(self):
        """Move selected question up in order within its set."""
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question to move.")
            return

        item = selected_items[0]
        if item.parent() is None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a question item (not a group).")
            return

        uniqueid = item.data(0, Qt.ItemDataRole.UserRole)
        if not uniqueid:
            return

        question = next((q for q in self.questions if q["uniqueid"] == uniqueid), None)
        if not question:
            return

        set_id = question.get("set_id")
        set_questions = [q for q in self.questions if q.get("set_id") == set_id]
        set_questions.sort(key=lambda x: x.get("order", 999999))

        current_index = next((i for i, q in enumerate(set_questions) if q["uniqueid"] == uniqueid), -1)
        if current_index <= 0:
            return  # Already at the top

        # Swap orders
        set_questions[current_index]["order"], set_questions[current_index - 1]["order"] = \
            set_questions[current_index - 1]["order"], set_questions[current_index]["order"]

        self.update_order_numbers()
        self.update_preview()
        # Reselect the item
        self.select_question_in_tree(uniqueid)

    def move_question_down(self):
        """Move selected question down in order within its set."""
        selected_items = self.tree.selectedItems()
        if not selected_items:
            QMessageBox.warning(self, "No Selection", "Please select a question to move.")
            return

        item = selected_items[0]
        if item.parent() is None:
            QMessageBox.warning(self, "Invalid Selection", "Please select a question item (not a group).")
            return

        uniqueid = item.data(0, Qt.ItemDataRole.UserRole)
        if not uniqueid:
            return

        question = next((q for q in self.questions if q["uniqueid"] == uniqueid), None)
        if not question:
            return

        set_id = question.get("set_id")
        set_questions = [q for q in self.questions if q.get("set_id") == set_id]
        set_questions.sort(key=lambda x: x.get("order", 999999))

        current_index = next((i for i, q in enumerate(set_questions) if q["uniqueid"] == uniqueid), -1)
        if current_index >= len(set_questions) - 1:
            return  # Already at the bottom

        # Swap orders
        set_questions[current_index]["order"], set_questions[current_index + 1]["order"] = \
            set_questions[current_index + 1]["order"], set_questions[current_index]["order"]

        self.update_order_numbers()
        self.update_preview()
        # Reselect the item
        self.select_question_in_tree(uniqueid)

    def select_question_in_tree(self, uniqueid: str):
        """Select a question item in the tree by uniqueid."""
        def find_item(parent_item):
            for i in range(parent_item.childCount()):
                child = parent_item.child(i)
                if child.data(0, Qt.ItemDataRole.UserRole) == uniqueid:
                    self.tree.setCurrentItem(child)
                    return True
                if child.childCount() > 0:
                    if find_item(child):
                        return True
            return False

        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            if find_item(root.child(i)):
                break

    def extract_and_save_images(self, html_text: str, question_uniqueid: str, images_dir: Path) -> str:
        """Extract images from HTML, save them, and return plain text with newline-separated image references."""
        if not html_text:
            return ""
        
        # Find all img tags in HTML
        img_pattern = r'<img[^>]*src=["\']([^"\']+)["\'][^>]*>'
        
        # Split HTML by image tags, processing segments
        parts = []
        last_end = 0
        
        for match in re.finditer(img_pattern, html_text):
            # Add text before this image
            before_html = html_text[last_end:match.start()]
            if before_html.strip():
                from PyQt6.QtGui import QTextDocument
                doc = QTextDocument()
                doc.setHtml(before_html)
                plain = doc.toPlainText().strip()
                if plain:
                    parts.append(plain)
            
            # Process the image
            src = match.group(1)
            image_path = None
            
            # Check if it's a data URI (base64)
            if src.startswith("data:image"):
                # Extract base64 data
                base64_match = re.search(r'base64,(.+)', src)
                if base64_match:
                    try:
                        image_data = base64.b64decode(base64_match.group(1))
                        # Generate unique filename
                        image_uniqueid = str(uuid.uuid4())
                        image_filename = f"{image_uniqueid}.png"
                        image_path_file = images_dir / image_filename
                        
                        # Save image
                        with open(image_path_file, "wb") as f:
                            f.write(image_data)
                        
                        image_path = f"images/{image_filename}"
                    except Exception:
                        pass  # Skip invalid images
            
            # If it's already a file path (images/uuid.png format), use it
            elif src.startswith("images/") and src.endswith(".png"):
                image_path = src
            
            # Add image reference
            if image_path:
                parts.append(image_path)
            
            last_end = match.end()
        
        # Add remaining text after last image
        if last_end < len(html_text):
            remaining_html = html_text[last_end:]
            if remaining_html.strip():
                from PyQt6.QtGui import QTextDocument
                doc = QTextDocument()
                doc.setHtml(remaining_html)
                plain = doc.toPlainText().strip()
                if plain:
                    parts.append(plain)
        
        # If no images found, just return plain text
        if not parts:
            from PyQt6.QtGui import QTextDocument
            doc = QTextDocument()
            doc.setHtml(html_text)
            return doc.toPlainText()
        
        return "\n".join(parts)

    def update_preview(self):
        """Update the preview tree with current questions grouped by set."""
        # Save expansion state of groups before clearing
        expansion_state = {}
        root = self.tree.invisibleRootItem()
        for i in range(root.childCount()):
            group_item = root.child(i)
            set_id = group_item.data(0, Qt.ItemDataRole.UserRole)
            if set_id:
                expansion_state[set_id] = group_item.isExpanded()
        
        self.tree.clear()

        # Group questions by set_id and sort within each set by order
        questions_by_set = {}
        for question in self.questions:
            set_id = question.get("set_id")
            if set_id not in questions_by_set:
                questions_by_set[set_id] = []
            questions_by_set[set_id].append(question)
        
        # Sort questions within each set by order
        for set_id in questions_by_set:
            questions_by_set[set_id] = sorted(questions_by_set[set_id], key=lambda x: x.get("order", 999999))

        # Create tree structure with groups
        root = self.tree.invisibleRootItem()
        for qset in self.question_sets:
            set_id = qset["id"]
            if set_id not in questions_by_set:
                continue
            
            # Create group item (parent)
            section_text = f" (Section {qset['section']})" if qset['section'] else ""
            group_text = f"{qset['label']}{section_text}"
            group_item = QTreeWidgetItem([group_text, "", "", "", ""])
            
            # Restore expansion state if it exists, otherwise expand by default
            if set_id in expansion_state:
                group_item.setExpanded(expansion_state[set_id])
            else:
                group_item.setExpanded(True)  # Expanded by default for new groups
            
            # Store set_id in group item for reference
            group_item.setData(0, Qt.ItemDataRole.UserRole, qset["id"])
            
            # Make group items droppable (so questions can be dropped into them)
            group_item.setFlags(
                group_item.flags() | 
                Qt.ItemFlag.ItemIsDropEnabled
            )
            
            root.addChild(group_item)

            # Add questions as children of this group
            for question in questions_by_set[set_id]:
                # Get plain text preview (strip HTML tags)
                text_html = question.get("text_body", "")
                text_preview = self.strip_html(text_html)[:80] + "..." if len(self.strip_html(text_html)) > 80 else self.strip_html(text_html)
                
                mark_html = question.get("mark_scheme", "")
                mark_preview = self.strip_html(mark_html)[:60] + "..." if len(self.strip_html(mark_html)) > 60 else self.strip_html(mark_html)

                # Create question item
                question_item = QTreeWidgetItem([
                    str(question.get("order", "")),  # Order
                    question["path"],  # Path
                    ANSWER_TYPES[question["answer_type"]],  # Answer Type
                    text_preview,  # Question Text Preview
                    mark_preview if question["mark_scheme"] else "(none)"  # Mark Scheme Preview
                ])
                
                # Store uniqueid in the item data
                question_item.setData(0, Qt.ItemDataRole.UserRole, question["uniqueid"])
                
                # Make question items draggable and droppable
                question_item.setFlags(
                    question_item.flags() | 
                    Qt.ItemFlag.ItemIsDragEnabled | 
                    Qt.ItemFlag.ItemIsDropEnabled
                )
                
                # Add to group
                group_item.addChild(question_item)

    def strip_html(self, html_text: str) -> str:
        """Strip HTML tags to get plain text preview."""
        from PyQt6.QtGui import QTextDocument
        doc = QTextDocument()
        doc.setHtml(html_text)
        return doc.toPlainText()

    def update_status(self):
        """Update the status bar."""
        count = len(self.questions)
        sets_count = len(self.question_sets)
        self.status_bar.showMessage(f"Ready | {sets_count} set(s) | {count} question(s) entered")

    def export_csv(self):
        """Export questions to CSV file."""
        if not self.questions:
            QMessageBox.warning(self, "No Data", "Please add at least one question before exporting.")
            return

        exam_name = self.exam_edit.text().strip()
        if not exam_name:
            QMessageBox.critical(self, "Error", "Please enter an exam name.")
            return
        
        # Validate exam name format (M or N followed by two digits)
        if not re.match(r'^[MN]\d{2}', exam_name):
            QMessageBox.warning(
                self, 
                "Exam Name Warning", 
                "Exam name should start with M or N followed by two digits (e.g., M22, N23).\n"
                "You can continue, but please verify the format is correct."
            )

        # Ask for save location with default filename based on exam name
        exam_name_clean = exam_name.replace(' ', '_')
        default_filename = f"{exam_name_clean}_questions.csv"
        filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save CSV File",
            default_filename,
            "CSV Files (*.csv);;All Files (*)"
        )

        if not filename:
            return

        try:
            # Create images directory next to the CSV file
            csv_path = Path(filename)
            images_dir = csv_path.parent / "images"
            images_dir.mkdir(exist_ok=True)
            
            with open(filename, "w", newline="", encoding="utf-8") as csvfile:
                writer = csv.DictWriter(
                    csvfile, 
                    fieldnames=CSV_HEADERS,
                    quoting=csv.QUOTE_ALL,  # Quote ALL fields to handle commas and special characters
                    doublequote=True  # Escape quotes by doubling them
                )
                writer.writeheader()

                # Export in order - sorted by set, then by order within each set
                # Group by set for better organization
                questions_by_set = {}
                for question in self.questions:
                    set_id = question.get("set_id")
                    if set_id not in questions_by_set:
                        questions_by_set[set_id] = []
                    questions_by_set[set_id].append(question)
                
                # Sort sets by label, then sort questions within each set by order
                sorted_sets = sorted(self.question_sets, key=lambda x: x.get("label", ""))
                
                for qset in sorted_sets:
                    set_id = qset["id"]
                    if set_id in questions_by_set:
                        set_questions = sorted(questions_by_set[set_id], key=lambda x: x.get("order", 999999))
                        for question in set_questions:
                            # Extract and save images, get updated HTML with image paths
                            text_body_html = question.get("text_body", "")
                            mark_scheme_html = question.get("mark_scheme", "")
                            
                            # Extract images and save them, get plain text with newline-separated image references
                            text_body_plain = self.extract_and_save_images(
                                text_body_html, question["uniqueid"], images_dir
                            )
                            mark_scheme_plain = self.extract_and_save_images(
                                mark_scheme_html, question["uniqueid"], images_dir
                            )
                            
                            writer.writerow({
                                "uniqueid": question["uniqueid"],
                                "path": question["path"],
                                "text_body": text_body_plain,
                                "answer_type": question["answer_type"],
                                "mark_scheme": mark_scheme_plain,
                                "needs_context": str(question["needs_context"]).lower(),
                                "exam": question["exam"],
                                "section": question["section"],
                                "topic": question["topic"],
                                "order": question.get("order", ""),
                                "marks": question.get("marks", ""),
                            })

            QMessageBox.information(
                self,
                "Success",
                f"CSV file and images exported successfully!\n\nCSV: {filename}\nImages folder: {images_dir}\n\n{len(self.questions)} question(s) exported."
            )
        except Exception as e:
            QMessageBox.critical(self, "Export Error", f"Failed to export CSV file:\n{str(e)}")


def main():
    """Main entry point for the application."""
    app = QApplication(sys.argv)
    
    # Set application style
    app.setStyle("Fusion")
    
    window = IBQuestionEntryApp()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
