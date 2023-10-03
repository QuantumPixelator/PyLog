import winrm
import threading
from PySide6.QtWidgets import QApplication, QMainWindow, QLabel, QLineEdit, QPushButton, QFrame, QVBoxLayout, QComboBox, QSlider, QTextEdit, QFileDialog, QGroupBox, QFormLayout, QHBoxLayout
from PySide6.QtCore import Qt, Signal, QObject
from PySide6.QtGui import QIcon

# Global variable to hold the winrm session
session = None

class LogUpdater(QObject):
    update_log_signal = Signal(str)

log_updater = LogUpdater()

def update_log(logs):
    log_text.setPlainText(logs)

def get_log_names():
    ps_script = "Get-EventLog -List | Select-Object -Property Log"
    result = session.run_ps(ps_script)
    return result.std_out.decode("utf-8").strip().split("\r\n")[2:]

def connect():
    global session
    server_address = server_entry.text()
    username = username_entry.text()
    password = password_entry.text()
    session = winrm.Session(f"http://{server_address}:5985/wsman", auth=(username, password), transport="ntlm")
    
    log_names = get_log_names()
    log_dropdown.clear()
    log_dropdown.addItems(log_names)

    connection_status_label.setText("Connected")
    connection_status_label.setStyleSheet("color: green;")
    connect_button.setText("Disconnect")
    connect_button.clicked.connect(disconnect)

def disconnect():
    global session
    session = None
    connection_status_label.setText("Disconnected")
    connection_status_label.setStyleSheet("color: red;")
    connect_button.setText("Connect")
    connect_button.clicked.connect(connect)

def fetch_logs():
    def run_fetch():
        selected_log = log_dropdown.currentText()
        log_type = type_dropdown.currentText()
        number_of_logs = log_amount_slider.value()
        ps_script = f"Get-EventLog -LogName {selected_log} -EntryType {log_type} -Newest {number_of_logs} | Format-List"
        result = session.run_ps(ps_script)
        logs = result.std_out.decode("utf-8")
        logs = logs.replace('\n', '\n' + '-' * 80 + '\n')

        log_updater.update_log_signal.emit(logs)

    threading.Thread(target=run_fetch, daemon=True).start()

def save_logs():
    filename, _ = QFileDialog.getSaveFileName(None, "Save Logs", "", "Log Files (*.log);;All Files (*)")
    if filename:
        with open(filename, 'w') as file:
            file.write(log_text.toPlainText())

def update_slider_label(value):
    log_amount_label.setText(f"Logs to retrieve: {value}")

app = QApplication([])

app.setStyleSheet("""
    QMainWindow {
        background-color: #34495e;
    }
    QLabel {
        font-size: 12px;
        color: #ecf0f1;
    }
    QLineEdit, QComboBox, QTextEdit, QSlider {
        background-color: #ecf0f1;
    }
    QPushButton {
        background-color: #3498db;
        color: #ecf0f1;
        padding: 5px;
        border-radius: 3px;
    }
    QGroupBox {
        border: 1px solid #ecf0f1;
        border-radius: 5px;
        margin-top: 1ex;
        color: #ecf0f1;
    }
    QGroupBox::title {
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 3px 0 3px;
    }
""")

window = QMainWindow()
window.setWindowTitle("PyLog - Windows Event Log Viewer")
window.setWindowIcon(QIcon("icon.ico"))
window.resize(600, 600)

central_widget = QFrame()
window.setCentralWidget(central_widget)

layout = QVBoxLayout()

connection_group = QGroupBox("Connection")
connection_layout = QFormLayout()
connection_status_label = QLabel("Disconnected")
connection_status_label.setStyleSheet("color: red;")
server_entry = QLineEdit()
server_entry.setPlaceholderText("Enter Server Address")
username_entry = QLineEdit()
username_entry.setPlaceholderText("Enter Username")
password_entry = QLineEdit()
password_entry.setPlaceholderText("Enter Password")
password_entry.setEchoMode(QLineEdit.Password)
connect_button = QPushButton("Connect")
connect_button.clicked.connect(connect)
connection_layout.addRow("Status:", connection_status_label)
connection_layout.addRow("Server:", server_entry)
connection_layout.addRow("Username:", username_entry)
connection_layout.addRow("Password:", password_entry)
connection_layout.addRow(connect_button)
connection_group.setLayout(connection_layout)

logs_group = QGroupBox("Logs")
logs_layout = QHBoxLayout()
log_dropdown = QComboBox()
font_metrics = log_dropdown.fontMetrics()
log_dropdown.setFixedWidth(font_metrics.horizontalAdvance('X') * 25)
type_dropdown = QComboBox()
type_dropdown.addItems(["Error", "Warning", "Information", "SuccessAudit", "FailureAudit"])
log_amount_slider = QSlider(Qt.Horizontal)
log_amount_slider.setRange(1, 50)
log_amount_label = QLabel("Logs to retrieve: 1")
log_amount_slider.valueChanged.connect(update_slider_label)
logs_layout.addWidget(log_dropdown)
logs_layout.addWidget(type_dropdown)
logs_layout.addWidget(log_amount_label)
logs_layout.addWidget(log_amount_slider)
logs_group.setLayout(logs_layout)

text_group = QGroupBox("Text")
text_layout = QVBoxLayout()
log_text = QTextEdit()
log_text.setPlaceholderText("Logs will be displayed here.")
fetch_button = QPushButton("Fetch Logs")
fetch_button.clicked.connect(fetch_logs)
save_button = QPushButton("Save")
save_button.clicked.connect(save_logs)
text_layout.addWidget(log_text)
text_layout.addWidget(fetch_button)
text_layout.addWidget(save_button)
text_group.setLayout(text_layout)

log_updater.update_log_signal.connect(update_log)

layout.addWidget(connection_group)
layout.addWidget(logs_group)
layout.addWidget(text_group)
central_widget.setLayout(layout)

window.show()
app.exec_()
