# gui.py
import sys
import os
import shlex
import subprocess
import traceback
import datetime
from PyQt6.QtCore import Qt, QObject, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QTextCursor
from PyQt6.QtWidgets import (
    QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,
    QTextEdit, QLabel, QFrame, QInputDialog
)
# Uncaught Exception Logger (module-level, not inside a class)
# Logs uncaught exceptions globally for the application.
LOG_PATH = r"C:\Users\NADLUROB\Desktop\Dash\log.txt"


def log_uncaught_exceptions(exc_type, exc_value, exc_tb):
    try:
        os.makedirs(os.path.dirname(LOG_PATH), exist_ok=True)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n[UNCAUGHT EXCEPTION] {}\n".format(datetime.datetime.now()))
            f.write(''.join(traceback.format_exception(exc_type, exc_value, exc_tb)))
            f.write("\n" + "-" * 60 + "\n")
    except (OSError, PermissionError):
        # If logging itself fails, fall back to default handler
        pass
    finally:
        # call default handler so the debugger / environment sees it too
        try:
            sys.__excepthook__(exc_type, exc_value, exc_tb)
        except (TypeError, RuntimeError, SystemExit):
            # Avoid raising from excepthook itself
            pass


sys.excepthook = log_uncaught_exceptions


# Worker that runs a command list safely and streams stdout
class ScriptRunnerWorker(QObject):
    output = pyqtSignal(str)
    finished = pyqtSignal()
    error = pyqtSignal(str)

    def __init__(self, cmd_list, cwd=None, log_path=LOG_PATH):
        super().__init__()
        # cmd_list must be a list (no shell=True). Example: [sys.executable, "-u", "script.py"]
        self.cmd_list = cmd_list
        self.cwd = cwd
        self.log_path = log_path

    def _open_log(self):
        if not self.log_path:
            return None
        try:
            os.makedirs(os.path.dirname(self.log_path), exist_ok=True)
            return open(self.log_path, "a", encoding="utf-8", buffering=1)
        except (OSError, PermissionError):
            return None

    def run(self):
        logf = self._open_log()
        try:
            # Ensure child Python prints crash info on native faults
            env = os.environ.copy()
            env.setdefault("PYTHONFAULTHANDLER", "1")

            proc = subprocess.Popen(
                self.cmd_list,
                cwd=self.cwd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,  # replaces universal_newlines=True
                encoding="utf-8",  # force UTF-8
                errors="replace"  # avoid crashes if bad bytes slip in
            )

            header = f"\n--- START {datetime.datetime.now():%Y-%m-%d %H:%M:%S} CMD: {' '.join(self.cmd_list)} ---\n"
            if logf:
                try:
                    logf.write(header)
                except (OSError, ValueError):
                    pass

            # streamlines to UI and log file
            for line in iter(proc.stdout.readline, ''):
                if not line:
                    break
                line = line.rstrip("\n")
                # Emit to UI thread (guard emitter errors)
                try:
                    self.output.emit(line)
                except (TypeError, RuntimeError):
                    # if the Qt signal cannot accept the item or thread issues occur, log and continue
                    if logf:
                        try:
                            logf.write(f"[EMIT ERROR] failed to emit line: {line}\n")
                        except (OSError, ValueError):
                            pass

                # Write to log if available
                if logf:
                    try:
                        logf.write(line + "\n")
                    except (OSError, ValueError):
                        pass

            # close stdout, wait for exit
            try:
                if proc.stdout:
                    proc.stdout.close()
            except (OSError, AttributeError):
                if logf:
                    try:
                        logf.write("[STDOUT CLOSE ERROR]\n")
                    except (OSError, ValueError):
                        pass

            try:
                ret = proc.wait()
            except (subprocess.TimeoutExpired, OSError) as e:
                # treat as subprocess error
                if logf:
                    try:
                        logf.write(f"[WAIT ERROR] {e}\n")
                    except (OSError, ValueError):
                        pass
                self.error.emit(f"Process wait failed: {e}")
                return

            trailer = f"--- END {datetime.datetime.now():%Y-%m-%d %H:%M:%S} EXIT CODE: {ret} ---\n"
            if logf:
                try:
                    logf.write(trailer)
                except (OSError, ValueError):
                    pass

            if ret != 0:
                # For native crashes, ret will be non-zero (Windows: 0xC0000409 as decimal -1073740791)
                self.error.emit(f"Process exited with code {ret}. See log: {self.log_path or 'n/a'}")

        except (FileNotFoundError, OSError, ValueError, subprocess.SubprocessError, PermissionError) as e:
            tb = traceback.format_exc()
            if logf:
                try:
                    logf.write(f"Exception in worker:\n{tb}\n")
                except (OSError, ValueError):
                    pass
            self.error.emit(f"Process failed: {e}. See log: {self.log_path or 'n/a'}")
        finally:
            if logf:
                try:
                    logf.flush()
                    logf.close()
                except (OSError, ValueError):
                    pass
            self.finished.emit()


# Simple terminal widget
class Terminal(QTextEdit):
    def __init__(self):
        super().__init__()
        self.setReadOnly(True)
        self.setFont(QFont("Courier New", 9))
        self.setStyleSheet("""
            QTextEdit {
                background-color: #1e1e1e;
                color: #c5c5c5;
                border: 2px solid #3399ff;
                border-radius: 6px;
            }
        """)
        # initial message
        self.append("> Live terminal ready.")

    def write_output(self, text):
        self.append(text)
        self.moveCursor(QTextCursor.MoveOperation.End)

    def clear_terminal(self):
        self.clear()
        self.append("> Terminal cleared.")


# Drag & drop button for PDFs
class DragDropButton(QPushButton):
    def __init__(self, name, terminal=None, parent=None):
        super().__init__(name, parent)
        self.setAcceptDrops(True)
        self.terminal = terminal
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setMinimumHeight(30)
        self.setStyleSheet("""
            QPushButton {
                background-color: #1a1a1a;
                border: 2px solid #3399ff;
                padding: 6px;
                border-radius: 8px;
                color: white;
            }
            QPushButton:hover {
                background-color: #222;
            }
        """)
        self._thread = None
        self._worker = None

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if url.toLocalFile().lower().endswith(".pdf"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                file_path = url.toLocalFile()
                if file_path.lower().endswith(".pdf"):
                    if self.terminal:
                        self.terminal.append(f"> Dropped file: {file_path}")
                    self.run_delegation_script_threaded(file_path)
            event.acceptProposedAction()
        else:
            event.ignore()

    def run_delegation_script_threaded(self, filepath):
        # Use the same python interpreter as this GUI to avoid environment mismatch.
        script_path = os.path.join(os.path.dirname(__file__), "process_delegation.py")
        cmd_list = [sys.executable, "-u", script_path, filepath]
        if self.terminal:
            self.terminal.append(f"> Running: {sys.executable} -u \"{script_path}\" \"{filepath}\"")

        self._thread = QThread()
        self._worker = ScriptRunnerWorker(cmd_list, cwd=os.path.dirname(__file__))
        self._worker.moveToThread(self._thread)

        self._thread.started.connect(self._worker.run)
        self._worker.output.connect(self.terminal.write_output)
        self._worker.error.connect(lambda e: self.terminal.append(f"[ERROR] {e}"))
        self._worker.finished.connect(self._thread.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self._thread.finished.connect(self._thread.deleteLater)

        self._thread.start()


# Custom title bar
class CustomTitleBar(QFrame):
    toggle_terminal = pyqtSignal()

    def __init__(self, parent):
        super().__init__(parent)
        self.offset = None
        self.toggle_btn = QPushButton("^")
        self.title = QLabel("PYTHON SCRIPT RUNNER")
        self.parent = parent
        self.setFixedHeight(40)
        self.setStyleSheet("background-color: #0c0c0c; border-bottom: 2px solid #3399ff;")
        self.init_ui()

    def init_ui(self):
        layout = QHBoxLayout()
        layout.setContentsMargins(10, 0, 10, 0)

        self.title.setFont(QFont("Segoe UI", 10))
        self.title.setStyleSheet("color: white;")
        self.title.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)

        self.toggle_btn.setFixedSize(30, 30)
        self.toggle_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.toggle_btn.setStyleSheet("""
            QPushButton {
                color: white;
                background-color: #1a1a1a;
                border: 2px solid #3399ff;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #3399ff;
            }
        """)
        self.toggle_btn.clicked.connect(self.toggle_terminal.emit)

        close_btn = QPushButton("✕")
        close_btn.setFixedSize(30, 30)
        close_btn.setStyleSheet("""
            QPushButton {
                color: white;
                background-color: #1a1a1a;
                border: 2px solid #3399ff;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #cc0000;
            }
        """)
        close_btn.clicked.connect(self.parent.close)

        layout.addWidget(self.title)
        layout.addStretch()
        layout.addWidget(self.toggle_btn)
        layout.addWidget(close_btn)
        self.setLayout(layout)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.offset = event.pos()

    def mouseMoveEvent(self, event):
        if hasattr(self, "offset"):
            self.parent.move(self.parent.pos() + event.pos() - self.offset)


# Main application class
class ScriptRunnerApp(QWidget):
    EDGE_MARGIN = 6

    def receive_logs_with_month(self):
        month, ok = QInputDialog.getText(self, "Select Month", "Enter month (e.g., January):")
        if ok and month.strip():
            # Pass the month as a parameter to the script
            self.run_command(f'powershell.exe ./Script2.ps1 -SelectedMonth "{month.strip()}"')

        else:
            self.terminal.append("[INFO]: Month selection cancelled.")
    def __init__(self):
        super().__init__()
        # Buttons map - prefer run_script for launching Python scripts so sys.executable is used.
        self.buttons = {
            "Send Logs": lambda: self.run_command(r'powershell.exe -File "C:\Users\NADLUROB\Documents\DelegationScript\Script1.ps1"'),
            "Receive Logs": self.receive_logs_with_month,
            "Inbox Sort": lambda: self.run_command("powershell.exe ./Script3.ps1"),
            "Delegations": None,
            # Use run_script to ensure same python interpreter is used and cwd is set.
            "Bulk Email": lambda: self.run_script(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Bulk Mail.py")),
            "Statements": lambda: self.run_script(os.path.join(os.path.dirname(os.path.abspath(__file__)), "Create Statements.py")),
            "Clear": self.clear_terminal,
        }



        self.buttons_container = QWidget()
        self.terminal = Terminal()
        self.terminal_label = None
        self.terminal_container = QWidget()
        self.title_bar = CustomTitleBar(self)
        self.setWindowFlags(Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint)
        self.setStyleSheet("background-color: #111; color: white;")
        self.setGeometry(0, 0, 500, 320)

        self.terminal_collapsed = False
        self._threads = []
        self._resizing = False
        self._resize_edge = None
        self._drag_pos = None
        self._start_geo = None

        self.init_ui()
        self.title_bar.toggle_terminal.connect(self.toggle_terminal_visibility)

    def init_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(2, 2, 2, 2)
        main_layout.setSpacing(6)

        main_layout.addWidget(self.title_bar)

        terminal_layout = QVBoxLayout(self.terminal_container)
        terminal_layout.setContentsMargins(0, 0, 0, 0)

        self.terminal_label = QLabel("LIVE TERMINAL")
        self.terminal_label.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.terminal_label.setStyleSheet("color: #88ccff;")

        self.terminal.setMinimumHeight(140)

        terminal_layout.addWidget(self.terminal_label)
        terminal_layout.addWidget(self.terminal)

        self.buttons_container.setFixedWidth(500)
        buttons_layout = QHBoxLayout(self.buttons_container)
        buttons_layout.setSpacing(6)
        buttons_layout.setContentsMargins(0, 0, 0, 0)

        for name, handler in self.buttons.items():
            btn = DragDropButton(name, self.terminal) if name == "Delegations" else QPushButton(name)
            btn.setFont(QFont("Segoe UI", 7))
            btn.setFixedWidth(66)
            btn.setMinimumHeight(28)
            if handler:
                btn.clicked.connect(handler)
            if name != "Delegations":
                btn.setStyleSheet("""
                    QPushButton {
                        background-color: #1a1a1a;
                        border: 2px solid #3399ff;
                        padding: 6px;
                        border-radius: 8px;
                        color: white;
                    }
                    QPushButton:hover {
                        background-color: #222;
                    }
                """)
            buttons_layout.addWidget(btn)

        main_layout.addWidget(self.buttons_container)
        main_layout.addWidget(self.terminal_container)

    def toggle_terminal_visibility(self):
        if self.terminal_collapsed:
            self.terminal_container.show()
            self.setMinimumSize(500, 320)
            self.setMaximumSize(500, 16777215)
            self.resize(500, 320)
        else:
            self.terminal_container.hide()
            self.setFixedSize(500, 100)
        self.terminal_collapsed = not self.terminal_collapsed

    # Centralized process launcher to reduce duplication and ensure correct environment.
    def run_process(self, cmd_list, cwd=None, display_cmd=None):
        disp = display_cmd or " ".join(shlex.quote(p) for p in cmd_list)
        self.terminal.append(f"> Running: {disp}")

        thread = QThread()
        worker = ScriptRunnerWorker(cmd_list, cwd=cwd)
        worker.moveToThread(thread)

        thread.started.connect(worker.run)
        worker.output.connect(self.terminal.write_output)
        worker.error.connect(lambda e: self.terminal.append(f"[ERROR]: {e}"))
        worker.finished.connect(thread.quit)
        worker.finished.connect(worker.deleteLater)
        thread.finished.connect(thread.deleteLater)

        thread.start()
        # track both thread and worker so we can keep them alive until finished.
        self._threads.append((thread, worker))
        # prune finished threads
        new_threads = []
        for t, w in self._threads:
            try:
                if t is not None and not t.isFinished():
                    new_threads.append((t, w))
            except RuntimeError:
                # Thread object invalid, ignore it
                pass
        self._threads = new_threads

    def run_command(self, command: str):
        # Use shlex for shell-like commands; these are not Python scripts.
        try:
            cmd_list = shlex.split(command)
        except ValueError as e:
            self.terminal.append(f"[ERROR]: invalid command string: {e}")
            return
        self.run_process(cmd_list, cwd=None, display_cmd=command)

    def run_script(self, script_full_path: str):
        # Important: run scripts with same interpreter as this GUI and unbuffered output.
        script_full_path = os.path.abspath(script_full_path)
        cmd_list = [sys.executable, "-u", script_full_path]
        cwd = os.path.dirname(script_full_path)
        display_cmd = f'"{sys.executable}" -u "{script_full_path}"'
        self.run_process(cmd_list, cwd=cwd, display_cmd=display_cmd)

    def clear_terminal(self):
        self.terminal.clear_terminal()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = ScriptRunnerApp()
    win.show()
    win.move(0, 0)
    sys.exit(app.exec())
