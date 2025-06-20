import os
import json
import shutil
import logging
import paramiko
from datetime import datetime
from tkinter import Tk, StringVar, filedialog, messagebox, Menu
from tkinter import ttk
from cryptography.fernet import Fernet
from PIL import Image, ImageTk
import urllib.request

CONFIG_FILE = "config.enc"
KEY_FILE = "config.key"

__version__ = "1.0.0"

class SFTPUploader:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF PO to EDI")
        self.root.geometry("750x550")
        self.root.configure(padx=20, pady=20, bg="#ffffff")
        self.config = None

        self.vars = {
            'SFTP_HOST': StringVar(),
            'SFTP_USERNAME': StringVar(),
            'SFTP_PASSWORD': StringVar(),
            'REMOTE_PATH': StringVar(),
            'ARCHIVE_FOLDER': StringVar(),
            'LOG_FILE': StringVar(value="sftp_transfer.log")
        }

        self.progress_var = StringVar(value="")

        self.setup_styles()
        self.build_gui()
        self.create_menu()
        self.load_or_configure()

    def setup_styles(self):
        style = ttk.Style()
        style.theme_use("default")

        style.configure("TFrame", background="#ffffff")
        style.configure("TLabel", background="#ffffff", foreground="#2E5A3B", font=("Segoe UI", 10))
        style.configure("TButton", background="#2E5A3B", foreground="#ffffff", font=("Segoe UI", 10), padding=6)
        style.map("TButton", background=[("active", "#3e704e")], foreground=[("active", "#ffffff")])

        style.configure("Treeview", rowheight=24, background="#ffffff", fieldbackground="#ffffff", font=("Segoe UI", 9))
        style.configure("Treeview.Heading", background="#2E5A3B", foreground="#ffffff", font=("Segoe UI", 10, "bold"))
        style.configure("TProgressbar", troughcolor="#dfeee4", background="#85C88A")

    def create_menu(self):
        menu = Menu(self.root)
        self.root.config(menu=menu)

        tools_menu = Menu(menu, tearoff=0)
        tools_menu.add_command(label="Check for Updates", command=self.check_for_updates)
        tools_menu.add_separator()
        tools_menu.add_command(label="Exit", command=self.root.quit)

        menu.add_cascade(label="Tools", menu=tools_menu)

    def build_gui(self):
        self.config_frame = ttk.Frame(self.root)
        self.upload_frame = ttk.Frame(self.root)

        self.field_map = [
            ("SFTP Host", "SFTP_HOST"),
            ("Username", "SFTP_USERNAME"),
            ("Password", "SFTP_PASSWORD"),
            ("Remote Path", "REMOTE_PATH"),
            ("Archive Folder", "ARCHIVE_FOLDER")
        ]

        for i, (label, var_key) in enumerate(self.field_map):
            ttk.Label(self.config_frame, text=label + ":").grid(row=i, column=0, sticky='e', pady=5)
            entry = ttk.Entry(
                self.config_frame,
                textvariable=self.vars[var_key],
                width=40,
                show='*' if label == "Password" else ''
            )
            entry.grid(row=i, column=1, pady=5)
            if label == "Archive Folder":
                ttk.Button(
                    self.config_frame, text="Browse", command=self.browse_folder
                ).grid(row=i, column=2, padx=5)

        ttk.Button(self.config_frame, text="Save & Continue", command=self.save_config).grid(
            row=len(self.field_map), columnspan=3, pady=15
        )

        upload_inner = ttk.Frame(self.upload_frame)
        upload_inner.pack(pady=10)

        try:
            logo_img = Image.open("ad_logo.jpg")
            logo_width = 180
            logo_ratio = logo_img.height / logo_img.width
            logo_height = int(logo_width * logo_ratio)
            logo_img = logo_img.resize((logo_width, logo_height), Image.Resampling.LANCZOS)
            self.upload_logo = ImageTk.PhotoImage(logo_img)
            ttk.Label(upload_inner, image=self.upload_logo, background="#ffffff").grid(row=0, column=0, rowspan=3, padx=30)
        except Exception as e:
            print("Logo load failed:", e)

        self.upload_btn = ttk.Button(upload_inner, text="Upload PDF Files", command=self.select_files)
        self.upload_btn.grid(row=0, column=1, pady=5)

        ttk.Button(upload_inner, text="Reconfigure Connection", command=self.reconfigure).grid(row=1, column=1, pady=5)
        ttk.Label(self.upload_frame, textvariable=self.progress_var).pack(pady=5)

        self.progress_bar = ttk.Progressbar(self.upload_frame, mode='determinate', length=700)
        self.progress_bar.pack(pady=5)

        self.result_table = ttk.Treeview(self.upload_frame, columns=("File", "Status", "Message"), show='headings')
        for col in ("File", "Status", "Message"):
            self.result_table.heading(col, text=col)
            if col == "File":
                self.result_table.column(col, width=340, anchor="w")
            elif col == "Status":
                self.result_table.column(col, width=100, anchor="center")
            else:
                self.result_table.column(col, width=260, anchor="w")
        self.result_table.pack(fill='both', expand=True, pady=10)

        self.result_table.tag_configure("Success", background="#d4edda")
        self.result_table.tag_configure("Failed", background="#f8d7da")
        self.result_table.tag_configure("Skipped", background="#fff3cd")

        self.button_frame = ttk.Frame(self.upload_frame)
        self.upload_more_btn = ttk.Button(self.button_frame, text="Upload More Files", command=self.select_files)
        self.cancel_btn = ttk.Button(self.button_frame, text="Exit", command=self.root.quit)
        self.upload_more_btn.pack(side="left", padx=10)
        self.cancel_btn.pack(side="left", padx=10)
        self.button_frame.pack(pady=10)
        self.button_frame.pack_forget()

    def browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.vars['ARCHIVE_FOLDER'].set(folder)

    def save_config(self):
        config = {k: v.get() for k, v in self.vars.items()}
        config['SFTP_PORT'] = 22

        try:
            transport = paramiko.Transport((config["SFTP_HOST"], config["SFTP_PORT"]))
            transport.connect(username=config["SFTP_USERNAME"], password=config["SFTP_PASSWORD"])
            sftp = paramiko.SFTPClient.from_transport(transport)

            sftp.listdir(config["REMOTE_PATH"])
            test_file = os.path.join(config["REMOTE_PATH"], "test_sftp.txt")
            with sftp.file(test_file, "w") as f:
                f.write("test")
            sftp.remove(test_file)

            sftp.close()
            transport.close()
        except Exception as e:
            messagebox.showerror("Connection Test Failed", f"Unable to validate SFTP connection:\n{e}")
            return

        archive_parent = os.path.dirname(config["ARCHIVE_FOLDER"])
        log_dir = os.path.join(archive_parent, "SFTP Logs")
        os.makedirs(log_dir, exist_ok=True)

        log_filename = f"sftp_transfer_{datetime.now().strftime('%Y%m%d')}.log"
        log_path = os.path.join(log_dir, log_filename)
        config["LOG_FILE"] = log_path

        key = self.get_key()
        fernet = Fernet(key)

        if os.path.exists(CONFIG_FILE):
            with open(CONFIG_FILE, 'rb') as f:
                backup_data = f.read()
            backup_name = f"{CONFIG_FILE}.backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            with open(backup_name, 'wb') as backup_file:
                backup_file.write(backup_data)

        encrypted = fernet.encrypt(json.dumps(config).encode())
        with open(CONFIG_FILE, 'wb') as f:
            f.write(encrypted)

        self.load_config()
        self.show_upload_frame()

    def load_or_configure(self):
        if os.path.exists(CONFIG_FILE):
            self.load_config()
            self.show_upload_frame()
        else:
            self.show_config_frame()

    def get_key(self):
        if not os.path.exists(KEY_FILE):
            with open(KEY_FILE, 'wb') as f:
                f.write(Fernet.generate_key())
        return open(KEY_FILE, 'rb').read()

    def load_config(self):
        key = self.get_key()
        fernet = Fernet(key)
        with open(CONFIG_FILE, 'rb') as f:
            self.config = json.loads(fernet.decrypt(f.read()).decode())
        for var_key in self.vars:
            self.vars[var_key].set(self.config.get(var_key, ""))

    def show_config_frame(self):
        self.upload_frame.pack_forget()
        self.config_frame.pack()

    def show_upload_frame(self):
        self.config_frame.pack_forget()
        self.upload_frame.pack(fill='both', expand=True)

    def reconfigure(self):
        self.show_config_frame()

    def select_files(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        pdf_files = [f for f in files if f.lower().endswith(".pdf")]
        if not pdf_files:
            messagebox.showinfo("No PDFs", "Please select PDF files only.")
            return
        self.upload_files(pdf_files)

    def upload_files(self, files):
        log_file = self.config.get("LOG_FILE", "")
        if not log_file:
            messagebox.showerror("Missing Log Path", "Log file path is not configured. Please reconfigure your connection.")
            return

        os.makedirs(os.path.dirname(log_file), exist_ok=True)
        logging.basicConfig(
            filename=log_file,
            filemode='a',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s'
        )

        self.result_table.delete(*self.result_table.get_children())
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(files)
        self.button_frame.pack_forget()
        self.upload_btn.config(state='disabled')

        try:
            transport = paramiko.Transport((self.config["SFTP_HOST"], self.config["SFTP_PORT"]))
            transport.connect(username=self.config["SFTP_USERNAME"], password=self.config["SFTP_PASSWORD"])
            sftp = paramiko.SFTPClient.from_transport(transport)
        except Exception as e:
            messagebox.showerror("SFTP Error", str(e))
            self.upload_btn.config(state='normal')
            return

        for idx, filepath in enumerate(files):
            filename = os.path.basename(filepath)
            remote_file = os.path.join(self.config["REMOTE_PATH"], filename)
            self.progress_var.set(f"Uploading: {filename}")
            self.progress_bar["value"] = idx
            self.root.update()
            overwrite = False

            logging.info(f"Preparing to upload: {filename}")

            try:
                sftp.stat(remote_file)
                if not messagebox.askyesno("Overwrite?", f"{filename} exists on the server. Overwrite?"):
                    self.log_result(filename, "Skipped", "User skipped overwrite", "Skipped")
                    logging.info(f"Skipped: {filename} (exists and user declined overwrite)")
                    continue
                overwrite = True
            except FileNotFoundError:
                pass

            try:
                sftp.put(filepath, remote_file)
                self.archive_file(filepath, overwrite)
                self.log_result(filename, "Success", "Uploaded", "Success")
                logging.info(f"Uploaded: {filename} | Remote Path: {remote_file} | Overwritten: {overwrite}")
            except Exception as e:
                self.log_result(filename, "Failed", str(e), "Failed")
                logging.error(f"Failed to upload: {filename} | Error: {str(e)}")

        sftp.close()
        transport.close()

        logging.info(f"Upload session complete. Total files: {len(files)}")
        self.progress_var.set("Upload Complete.")
        self.progress_bar["value"] = len(files)
        self.upload_btn.config(state='normal')
        self.button_frame.pack()

    def archive_file(self, filepath, overwrite):
        os.makedirs(self.config["ARCHIVE_FOLDER"], exist_ok=True)
        filename = os.path.basename(filepath)
        if overwrite:
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            name, ext = os.path.splitext(filename)
            filename = f"{name}_{ts}{ext}"
        shutil.move(filepath, os.path.join(self.config["ARCHIVE_FOLDER"], filename))

    def log_result(self, filename, status, msg, tag):
        self.result_table.insert('', 'end', values=(filename, status, msg), tags=(tag,))

    def check_for_updates(self):
        try:
            with urllib.request.urlopen(
                    "https://raw.githubusercontent.com/nobies123/po2edi/master/version.txt") as response:
                latest_version = response.read().decode().strip()

            if latest_version != __version__:
                if messagebox.askyesno("Update Available",
                                       f"A newer version ({latest_version}) is available. Open GitHub?"):
                    os.system("start https://github.com/nobies123/po2edi/releases")
            else:
                messagebox.showinfo("Up to Date", "You are using the latest version.")
        except Exception as e:
            messagebox.showerror("Could not check for updates", f"{e}")


if __name__ == "__main__":
    root = Tk()
    try:
        root.iconbitmap("ad_logo.ico")
    except:
        pass
    app = SFTPUploader(root)
    root.mainloop()
