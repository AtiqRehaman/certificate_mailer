import os
import ssl
import smtplib
import threading
from pathlib import Path
from email.message import EmailMessage
from tkinter import (
    Tk, Label, Button, Entry, StringVar, Text, END, filedialog,
    messagebox, DoubleVar, Spinbox, Frame
)

import pandas as pd
from PIL import Image, ImageDraw, ImageFont
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


APP_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = APP_DIR / "generated_certificates"
OUTPUT_DIR.mkdir(exist_ok=True)

# Coordinates calibrated for the supplied 1450x1000 certificate.
# They scale automatically if the Canva export has a different resolution.
DEFAULT_NAME_X = 725
DEFAULT_NAME_Y = 460
DEFAULT_FONT_SIZE = 46


class CertificateMailer:
    def __init__(self, root):
        self.root = root
        root.title("Bulk Certificate Mailer")
        root.geometry("900x760")

        self.template = None
        self.data = None

        self.template_var = StringVar(value="No template selected")
        self.excel_var = StringVar(value="No Excel/CSV selected")
        self.sender_var = StringVar()
        self.password_var = StringVar()
        self.subject_var = StringVar(value="Certificate of Achievement")
        self.font_var = StringVar(value="")
        self.x_var = StringVar(value=str(DEFAULT_NAME_X))
        self.y_var = StringVar(value=str(DEFAULT_NAME_Y))
        self.size_var = StringVar(value=str(DEFAULT_FONT_SIZE))

        Label(root, text="Bulk Certificate Mailer",
              font=("Arial", 22, "bold")).pack(pady=(15, 4))
        Label(root, text="Generate personalized certificates and email them individually.",
              font=("Arial", 11)).pack(pady=(0, 15))

        top = Frame(root)
        top.pack(fill="x", padx=20)

        Button(top, text="Choose Certificate PNG", command=self.choose_template).grid(
            row=0, column=0, padx=5, pady=5, sticky="w")
        Label(top, textvariable=self.template_var, anchor="w").grid(
            row=0, column=1, sticky="we")
        Button(top, text="Choose Excel / CSV", command=self.choose_data).grid(
            row=1, column=0, padx=5, pady=5, sticky="w")
        Label(top, textvariable=self.excel_var, anchor="w").grid(
            row=1, column=1, sticky="we")
        top.columnconfigure(1, weight=1)

        settings = Frame(root, relief="groove", borderwidth=1)
        settings.pack(fill="x", padx=20, pady=12)

        Label(settings, text="Name X").grid(row=0, column=0, padx=8, pady=8)
        Entry(settings, textvariable=self.x_var, width=10).grid(row=0, column=1)
        Label(settings, text="Name Y").grid(row=0, column=2, padx=8)
        Entry(settings, textvariable=self.y_var, width=10).grid(row=0, column=3)
        Label(settings, text="Font size").grid(row=0, column=4, padx=8)
        Entry(settings, textvariable=self.size_var, width=10).grid(row=0, column=5)

        Label(settings, text="Font file (optional)").grid(row=1, column=0, padx=8, pady=8)
        Entry(settings, textvariable=self.font_var, width=45).grid(row=1, column=1, columnspan=4, sticky="we")
        Button(settings, text="Browse", command=self.choose_font).grid(row=1, column=5, padx=8)

        email = Frame(root, relief="groove", borderwidth=1)
        email.pack(fill="x", padx=20, pady=5)

        Label(email, text="Gmail address").grid(row=0, column=0, padx=8, pady=8)
        Entry(email, textvariable=self.sender_var, width=35).grid(row=0, column=1, sticky="w")
        Label(email, text="Gmail App Password").grid(row=0, column=2, padx=8)
        Entry(email, textvariable=self.password_var, width=25, show="*").grid(row=0, column=3, sticky="w")

        Label(email, text="Email subject").grid(row=1, column=0, padx=8, pady=8)
        Entry(email, textvariable=self.subject_var, width=70).grid(row=1, column=1, columnspan=3, sticky="we")

        Label(root, text="Email body (use {name} for personalization):",
              font=("Arial", 10, "bold")).pack(anchor="w", padx=20, pady=(12, 3))
        self.body = Text(root, height=8, wrap="word")
        self.body.pack(fill="x", padx=20)
        self.body.insert("1.0",
            "Dear {name},\n\n"
            "Thank you for participating in the Technical Quiz Competition. We truly appreciate your enthusiasm and effort. Please find your certificate attached.\n"
            "Regards,\n"
            "Association of Computer Geeks (ACG)\n"
            "LBRCE")

        buttons = Frame(root)
        buttons.pack(pady=12)
        Button(buttons, text="Preview First Certificate", command=self.preview,
               width=25).grid(row=0, column=0, padx=8)
        Button(buttons, text="Generate All Certificates", command=self.generate_all,
               width=25).grid(row=0, column=1, padx=8)
        Button(buttons, text="Generate + Send All", command=self.send_all,
               width=25).grid(row=0, column=2, padx=8)

        self.status = StringVar(value="Ready")
        Label(root, textvariable=self.status, anchor="w").pack(fill="x", padx=20)

        Label(root, text="Expected columns: Name, Email. Optional: Certificate_ID",
              font=("Arial", 9)).pack(pady=4)

    def choose_template(self):
        path = filedialog.askopenfilename(
            title="Select certificate template",
            filetypes=[("Images", "*.png *.jpg *.jpeg")])
        if path:
            self.template = Path(path)
            self.template_var.set(str(self.template))

    def choose_data(self):
        path = filedialog.askopenfilename(
            title="Select participant list",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv")])
        if not path:
            return
        try:
            if path.lower().endswith(".csv"):
                self.data = pd.read_csv(path)
            else:
                self.data = pd.read_excel(path)
            required = {"Name", "Email"}
            missing = required - set(self.data.columns)
            if missing:
                raise ValueError(f"Missing columns: {', '.join(sorted(missing))}")
            self.excel_var.set(f"{path}  ({len(self.data)} recipients)")
            self.status.set(f"Loaded {len(self.data)} recipients.")
        except Exception as e:
            messagebox.showerror("Data error", str(e))

    def choose_font(self):
        path = filedialog.askopenfilename(
            title="Select font",
            filetypes=[("Font files", "*.ttf *.otf")])
        if path:
            self.font_var.set(path)

    def settings(self):
        return (
            float(self.x_var.get()),
            float(self.y_var.get()),
            int(self.size_var.get())
        )

    def font(self, size):
        path = self.font_var.get().strip()
        if path and Path(path).exists():
            return ImageFont.truetype(path, size)
        # Good general fallback for certificate-style names.
        candidates = [
            "/usr/share/fonts/truetype/liberation2/LiberationSerif-Regular.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSerif.ttf",
            "C:/Windows/Fonts/georgia.ttf",
            "C:/Windows/Fonts/times.ttf",
        ]
        for p in candidates:
            if Path(p).exists():
                return ImageFont.truetype(p, size)
        return ImageFont.load_default()

    def make_certificate(self, name, cert_id=None):
        if not self.template:
            raise ValueError("Choose a certificate template first.")

        image = Image.open(self.template).convert("RGB")
        base_w, base_h = 1450, 1000
        sx, sy = image.width / base_w, image.height / base_h
        x, y, size = self.settings()
        x, y, size = x * sx, y * sy, int(size * ((sx + sy) / 2))

        draw = ImageDraw.Draw(image)
        fnt = self.font(size)

        # Center the participant name at the configured X coordinate.
        bbox = draw.textbbox((0, 0), str(name), font=fnt)
        text_w = bbox[2] - bbox[0]
        draw.text((x - text_w / 2, y), str(name), font=fnt, fill=(15, 15, 15))

        if cert_id is not None and str(cert_id).strip():
            # Optional ID is placed unobtrusively near the lower-right corner.
            id_font = self.font(max(16, int(size * 0.34)))
            text = f"Certificate ID: {cert_id}"
            draw.text((image.width - 360, image.height - 55), text,
                      font=id_font, fill=(40, 40, 40))

        safe = "".join(c for c in str(name) if c.isalnum() or c in " _-").strip()
        png_path = OUTPUT_DIR / f"Certificate_{safe}.png"
        pdf_path = OUTPUT_DIR / f"Certificate_{safe}.pdf"
        image.save(png_path, "PNG")

        # Put the PNG into a PDF at its native aspect ratio.
        c = canvas.Canvas(str(pdf_path), pagesize=(image.width, image.height))
        c.drawImage(ImageReader(str(png_path)), 0, 0,
                    width=image.width, height=image.height)
        c.save()
        return pdf_path

    def preview(self):
        if self.data is None or self.data.empty:
            messagebox.showwarning("Preview", "Load the participant Excel/CSV first.")
            return
        try:
            row = self.data.iloc[0]
            pdf = self.make_certificate(row["Name"], row.get("Certificate_ID"))
            # Open the generated PDF using the OS default application.
            os.startfile(str(pdf)) if os.name == "nt" else os.system(f'xdg-open "{pdf}" >/dev/null 2>&1')
            self.status.set(f"Preview generated: {pdf.name}")
        except Exception as e:
            messagebox.showerror("Preview error", str(e))

    def generate_all(self):
        if self.data is None:
            messagebox.showwarning("Generate", "Load the participant Excel/CSV first.")
            return
        try:
            for _, row in self.data.iterrows():
                self.make_certificate(row["Name"], row.get("Certificate_ID"))
            self.status.set(f"Generated {len(self.data)} certificates in {OUTPUT_DIR}")
            messagebox.showinfo("Complete", f"Generated {len(self.data)} certificates.")
        except Exception as e:
            messagebox.showerror("Generation error", str(e))

    def send_all(self):
        if self.data is None:
            messagebox.showwarning("Send", "Load the participant Excel/CSV first.")
            return
        sender = self.sender_var.get().strip()
        password = self.password_var.get().strip()
        if not sender or not password:
            messagebox.showwarning(
                "Gmail setup",
                "Enter your Gmail address and a Gmail App Password.\n\n"
                "Do NOT use your normal Gmail password.")
            return

        if not messagebox.askyesno(
            "Confirm bulk send",
            f"Send personalized certificates to {len(self.data)} recipients?"):
            return

        threading.Thread(target=self._send_worker, daemon=True).start()

    def _send_worker(self):
        try:
            self.status.set("Generating certificates and connecting to Gmail...")
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
                server.login(self.sender_var.get().strip(), self.password_var.get().strip())

                body_template = self.body.get("1.0", END).strip()
                total = len(self.data)
                sent = 0

                for i, (_, row) in enumerate(self.data.iterrows(), start=1):
                    name = str(row["Name"]).strip()
                    recipient = str(row["Email"]).strip()
                    cert_id = row.get("Certificate_ID")

                    pdf = self.make_certificate(name, cert_id)

                    msg = EmailMessage()
                    msg["Subject"] = self.subject_var.get()
                    msg["From"] = self.sender_var.get().strip()
                    msg["To"] = recipient
                    msg.set_content(body_template.replace("{name}", name))

                    with open(pdf, "rb") as f:
                        msg.add_attachment(
                            f.read(),
                            maintype="application",
                            subtype="pdf",
                            filename=pdf.name
                        )

                    server.send_message(msg)
                    sent += 1
                    self.status.set(f"Sent {sent}/{total}: {name}")

            self.status.set(f"Finished. Successfully sent {sent}/{total}.")
            messagebox.showinfo("Bulk send complete", f"Sent {sent}/{total} certificates.")
        except Exception as e:
            self.status.set("Sending stopped because of an error.")
            messagebox.showerror("Email error", str(e))


if __name__ == "__main__":
    root = Tk()
    CertificateMailer(root)
    root.mainloop()
