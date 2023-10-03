import winrm
import threading
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import font


#########################################################
#                       FUNCTIONS                       #
#########################################################


def change_widget_font(widget, new_font):
    try:
        widget.config(font=new_font)
    except tk.TclError:
        pass

    for child in widget.winfo_children():
        change_widget_font(child, new_font)


def apply_font():
    selected_font = font_dropdown.get()
    selected_size = int(size_scale.get())
    new_font = f"{selected_font} {selected_size}"

    # Apply globally to new widgets
    window.option_add("*Font", new_font)

    # Apply to ttk widgets
    style = ttk.Style(window)
    style.configure(".", font=new_font)

    # Apply to all existing non-ttk widgets
    change_widget_font(window, new_font)


def connect():
    global session
    server_address = server_entry.get()
    username = username_entry.get()
    password = password_entry.get()
    session = winrm.Session(
        f"http://{server_address}:5985/wsman",
        auth=(username, password),
        transport="ntlm",
    )
    log_dropdown["values"] = get_log_names()
    connection_status_label.config(text="Connected".ljust(35), fg="green")
    connect_button.config(text="Disconnect", command=disconnect)


def disconnect():
    global session
    session = None
    connection_status_label.config(text="Disconnected".ljust(35), fg="red")
    connect_button.config(text="Connect", command=connect)
    log_dropdown["values"] = []
    log_dropdown.set("Select Log")


def get_log_names():
    ps_script = "Get-EventLog -List | Select-Object -Property Log"
    result = session.run_ps(ps_script)
    return result.std_out.decode("utf-8").strip().split("\r\n")[2:]


def fetch_logs():
    def run_fetch():
        connection_status_label.config(text="Retrieving logs...".ljust(35), fg="yellow")
        selected_log = log_dropdown.get()
        log_type = type_dropdown.get()
        number_of_logs = int(num_logs_entry.get())
        ps_script = f"""
        Get-EventLog -LogName {selected_log} -EntryType {log_type} -Newest {number_of_logs} | Format-List
        """
        result = session.run_ps(ps_script)
        log_text.delete(1.0, tk.END)
        log_text.insert(tk.END, result.std_out.decode("utf-8"))
        connection_status_label.config(text="Connected - Logs retrieved".ljust(35), fg="green")

    threading.Thread(target=run_fetch, daemon=True).start()


def save_logs():
    filename = filedialog.asksaveasfilename(
        defaultextension=".log",
        filetypes=[("Log Files", "*.log"), ("All Files", "*.*")],
    )
    if filename:
        with open(filename, "w") as file:
            file.write(log_text.get(1.0, tk.END))


session = None


#########################################################
#                            GUI                        #
#########################################################

# Create the main window with 800x600 dimensions
window = tk.Tk()
window.geometry("800x600")
window.title("PyLog - Windows Event Log Viewer")

# Create the frames
left_frame = tk.Frame(
    window, bg="#1A1A1A", width=300, height=600, borderwidth=1, relief="ridge"
)
left_frame.grid(row=0, column=0, sticky="nsew", padx=1, pady=1)

middle_frame = tk.Frame(
    window, bg="#1A1A1A", width=200, height=600, borderwidth=1, relief="ridge"
)
middle_frame.grid(row=0, column=1, sticky="nsew", padx=1, pady=1)

right_frame = tk.Frame(
    window, bg="#1A1A1A", width=230, height=600, borderwidth=1, relief="ridge"
)
right_frame.grid(row=0, column=2, sticky="nsew", padx=1, pady=1)

# Configure the columns to adjust their sizes as needed
# window.grid_columnconfigure(0, weight=2)
window.grid_columnconfigure(1, weight=1)
window.grid_columnconfigure(2, weight=2)

# Labels for column 1 (server stuff)
connection_status_label = tk.Label(
    left_frame, text="Disconnected".ljust(35), bg="#1A1A1A", fg="red", padx=5)
connection_status_label.grid(row=5, column=0, sticky="w")

server_label = tk.Label(left_frame, text="Server:", bg="#1A1A1A", fg="white", padx=5)
server_label.grid(row=0, column=0, sticky="w")

username_label = tk.Label(
    left_frame, text="Username:", bg="#1A1A1A", fg="white", padx=5
)
username_label.grid(row=1, column=0, sticky="w")

password_label = tk.Label(
    left_frame, text="Password:", bg="#1A1A1A", fg="white", padx=5
)
password_label.grid(row=2, column=0, sticky="w")

# Entry boxes and button for column 2 (server stuff)
server_entry = tk.Entry(left_frame, bg="white", fg="black")
server_entry.grid(row=0, column=1, sticky="w")
server_entry.insert(0, "Enter Server Address")

username_entry = tk.Entry(left_frame, bg="white", fg="black")
username_entry.grid(row=1, column=1, sticky="w")
username_entry.insert(0, "Enter Username")

password_entry = tk.Entry(left_frame, bg="white", fg="black", show="*")
password_entry.grid(row=2, column=1, sticky="w")

connect_button = tk.Button(
    left_frame, text="Connect", bg="#4CAF50", fg="white", command=connect
)
connect_button.grid(row=3, column=1, sticky="e")

# Comboboxes for column 3 (logs stuff)
log_dropdown = ttk.Combobox(middle_frame)
log_dropdown.grid(row=0, column=0, sticky="ew", pady=5)

type_dropdown = ttk.Combobox(
    middle_frame,
    values=["Error", "Warning", "Information", "SuccessAudit", "FailureAudit"],
)
type_dropdown.grid(row=1, column=0, sticky="ew", pady=5)

# Label for Amount to Fetch
amount_label = tk.Label(
    middle_frame, text="Amount to fetch:", bg="#1A1A1A", fg="white", padx=5
)
amount_label.grid(row=2, column=0, sticky="w")

# Scale for selecting amount
num_logs_entry = tk.Scale(
    middle_frame,
    from_=1,
    to=100,
    orient="horizontal",
    bg="#1A1A1A",
    fg="white",
    sliderlength=15,
)
num_logs_entry.grid(row=3, column=0, sticky="ew")

# Labels for column 4 (fonts stuff)
font_label = tk.Label(right_frame, text="Font:", bg="#1A1A1A", fg="white", padx=5)
font_label.grid(row=0, column=0, sticky="w")

size_label = tk.Label(right_frame, text="Size:", bg="#1A1A1A", fg="white", padx=5)
size_label.grid(row=1, column=0, sticky="w")

# Combobox, Scale, and Button for column 5 (fonts stuff)
font_dropdown = ttk.Combobox(right_frame, values=font.families())
font_dropdown.grid(row=0, column=1, sticky="ew", pady=5)

size_scale = tk.Scale(
    right_frame,
    from_=1,
    to=42,
    orient="horizontal",
    bg="#1A1A1A",
    fg="white",
    sliderlength=15,
)
size_scale.grid(row=1, column=1, sticky="ew")

set_font_button = tk.Button(
    right_frame, text="Set Font", bg="#4CAF50", fg="white", command=apply_font
)
set_font_button.grid(row=2, column=1, sticky="e")

# Frame to hold the save and fetch buttons
buttons_frame = tk.Frame(window, bg="#1A1A1A")
buttons_frame.grid(row=0, column=0, columnspan=3, padx=10, pady=10, sticky="e")

# Save button
save_button = tk.Button(
    buttons_frame, text="Save", command=save_logs, bg="#1A1A1A", fg="white"
)
save_button.grid(row=0, column=1, sticky="e")

# Fetch button
fetch_button = tk.Button(
    buttons_frame, text="Fetch Logs", command=fetch_logs, bg="#1A1A1A", fg="white"
)
fetch_button.grid(row=1, column=1, sticky="e")

# Textbox to display logs with vertical scrollbar
log_text_frame = tk.Frame(window, bg="#1A1A1A")
log_text_frame.grid(row=2, column=0, columnspan=3, padx=10, pady=10, sticky="nsew")
log_text_frame.grid_columnconfigure(0, weight=1)
log_text_frame.grid_rowconfigure(0, weight=1)
log_text = tk.Text(log_text_frame, wrap=tk.WORD, bg="black", fg="white")
log_text.grid(row=0, column=0, sticky="nsew")
scrollbar = tk.Scrollbar(log_text_frame, command=log_text.yview)
scrollbar.grid(row=0, column=1, sticky="ns")
log_text.config(yscrollcommand=scrollbar.set)

# Configure the main window to resize the text area
window.grid_rowconfigure(2, weight=1)
window.grid_columnconfigure(0, weight=1)


#########################################################
#                         MAIN LOOP                     #
#########################################################

window.mainloop()
