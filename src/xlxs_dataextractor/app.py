from __future__ import annotations

import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QApplication,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListWidget,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

from .exporter import write_results
from .extractor import ExtractionResult, extract_workbook
from .paths import default_output_path


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("XLXS Data Extractor")
        self.resize(860, 620)
        self.files: list[Path] = []
        self.existing_extract: Path | None = None

        self.file_list = QListWidget()
        self.file_list.setAcceptDrops(False)

        self.existing_extract_label = QLabel("No existing extract loaded.")
        self.existing_extract_label.setWordWrap(True)

        self.status = QTextEdit()
        self.status.setReadOnly(True)
        self.status.setPlaceholderText("Processing details will appear here.")

        select_button = QPushButton("Select XLSX Files")
        select_button.clicked.connect(self.select_files)

        existing_button = QPushButton("Load Existing Extract")
        existing_button.clicked.connect(self.select_existing_extract)

        clear_existing_button = QPushButton("Clear Existing Extract")
        clear_existing_button.clicked.connect(self.clear_existing_extract)

        clear_button = QPushButton("Clear")
        clear_button.clicked.connect(self.clear_files)

        process_button = QPushButton("Extract And Save XLSX")
        process_button.clicked.connect(self.process_files)

        button_row = QHBoxLayout()
        button_row.addWidget(select_button)
        button_row.addWidget(existing_button)
        button_row.addWidget(clear_existing_button)
        button_row.addWidget(clear_button)
        button_row.addStretch()
        button_row.addWidget(process_button)

        layout = QVBoxLayout()
        title = QLabel("Offline Discharge Summary XLSX Extractor")
        title.setStyleSheet("font-size: 20px; font-weight: 700;")
        layout.addWidget(title)
        layout.addWidget(
            QLabel(
                "Select one or more Excel files. Optionally load an existing extracted workbook "
                "and new rows will be appended underneath."
            )
        )
        layout.addLayout(button_row)
        layout.addWidget(QLabel("Existing extract (optional):"))
        layout.addWidget(self.existing_extract_label)
        layout.addWidget(QLabel("Selected files:"))
        layout.addWidget(self.file_list, stretch=2)
        layout.addWidget(QLabel("Status:"))
        layout.addWidget(self.status, stretch=3)

        root = QWidget()
        root.setLayout(layout)
        self.setCentralWidget(root)

    def select_files(self) -> None:
        filenames, _ = QFileDialog.getOpenFileNames(
            self,
            "Select XLSX files",
            str(Path.home()),
            "Excel files (*.xlsx)",
        )
        if not filenames:
            return

        known = {path.resolve() for path in self.files}
        for filename in filenames:
            path = Path(filename)
            if path.resolve() not in known:
                self.files.append(path)
                known.add(path.resolve())
                self.file_list.addItem(str(path))

        self._log(f"Added {len(filenames)} file(s).")

    def select_existing_extract(self) -> None:
        filename, _ = QFileDialog.getOpenFileName(
            self,
            "Load existing extracted workbook",
            str(default_output_path().parent),
            "Excel files (*.xlsx)",
        )
        if not filename:
            return

        self.existing_extract = Path(filename)
        self.existing_extract_label.setText(str(self.existing_extract))
        self._log(f"Loaded existing extract: {self.existing_extract.name}")

    def clear_existing_extract(self) -> None:
        self.existing_extract = None
        self.existing_extract_label.setText("No existing extract loaded.")
        self._log("Cleared existing extract.")

    def clear_files(self) -> None:
        self.files.clear()
        self.file_list.clear()
        self.status.clear()

    def process_files(self) -> None:
        if not self.files:
            QMessageBox.warning(self, "No files selected", "Please select at least one .xlsx file.")
            return

        output_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save extracted data",
            str(self.existing_extract or default_output_path()),
            "Excel files (*.xlsx)",
        )
        if not output_path:
            return

        if not output_path.lower().endswith(".xlsx"):
            output_path += ".xlsx"

        self.status.clear()
        results: list[ExtractionResult] = []
        failures: list[str] = []

        for file_path in self.files:
            try:
                result = extract_workbook(file_path)
                results.append(result)
                if result.missing_fields:
                    self._log(f"{file_path.name}: processed, missing {', '.join(result.missing_fields)}")
                else:
                    self._log(f"{file_path.name}: processed successfully")
            except Exception as exc:  # noqa: BLE001 - UI should show per-file failures and continue.
                failures.append(f"{file_path.name}: {exc}")
                self._log(f"{file_path.name}: failed - {exc}")

        if not results:
            QMessageBox.critical(self, "Extraction failed", "No files were extracted successfully.")
            return

        try:
            write_results(results, output_path, existing_path=self.existing_extract)
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, "Save failed", f"Could not save output file:\n{exc}")
            return

        appended_note = ""
        if self.existing_extract:
            appended_note = f"\nAppended to existing data from:\n{self.existing_extract.name}"

        summary = f"Saved {len(results)} new row(s) to:\n{output_path}{appended_note}"
        if failures:
            summary += f"\n\n{len(failures)} file(s) failed. See status area for details."
        QMessageBox.information(self, "Done", summary)
        self._log(summary)

    def _log(self, message: str) -> None:
        self.status.append(message)
        self.status.verticalScrollBar().setValue(self.status.verticalScrollBar().maximum())


def main() -> int:
    app = QApplication(sys.argv)
    app.setAttribute(Qt.ApplicationAttribute.AA_DontShowIconsInMenus, False)
    window = MainWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
