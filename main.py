import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import time
import base64
from difflib import SequenceMatcher
from PIL import Image, ImageTk
from docx import Document
import easyocr
import requests

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


def _similarity(a, b):
    return round(SequenceMatcher(None, a.lower(), b.lower()).ratio() * 100, 1)


class ComposerApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Composer by Imtiaz Ahmad")
        self.geometry("1200x700")
        self.minsize(900, 600)
        self.resizable(True, True)

        self.reader = None
        self.current_image_path = None
        self._compare_mode = False

        self._build_ui()
        self._load_model_async()

    # ── UI construction ─────────────────────────────────────────────────────

    def _build_ui(self):
        header = ctk.CTkFrame(self, height=70, corner_radius=0, fg_color="#1a1a2e")
        header.pack(fill="x")
        header.pack_propagate(False)

        ctk.CTkLabel(
            header,
            text="Composer  by Imtiaz Ahmad",
            font=ctk.CTkFont(family="Segoe UI", size=22, weight="bold"),
            text_color="#e0e0ff",
        ).pack(side="left", padx=25, pady=15)

        self.status_badge = ctk.CTkLabel(
            header,
            text="⏳  Loading OCR model...",
            font=ctk.CTkFont(size=12),
            text_color="orange",
        )
        self.status_badge.pack(side="right", padx=25)

        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=18, pady=14)

        # ── Left panel ──────────────────────────────────────────────────────
        left = ctk.CTkFrame(body, width=380)
        left.pack(side="left", fill="both", expand=False, padx=(0, 10))
        left.pack_propagate(False)

        ctk.CTkLabel(
            left, text="Paper Image", font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(pady=(14, 8))

        self.drop_zone = ctk.CTkFrame(left, fg_color="#1e1e2e", height=230, corner_radius=12)
        self.drop_zone.pack(padx=14, fill="x")
        self.drop_zone.pack_propagate(False)

        self.img_hint = ctk.CTkLabel(
            self.drop_zone,
            text="📄\n\nClick  'Select Image'  below\nto upload a scanned paper",
            font=ctk.CTkFont(size=13),
            text_color="gray55",
            justify="center",
        )
        self.img_hint.pack(expand=True)

        ctk.CTkButton(
            left, text="Select Image", height=42,
            font=ctk.CTkFont(size=14), command=self._pick_image,
        ).pack(padx=14, pady=10, fill="x")

        # Language selector
        ctk.CTkLabel(left, text="Language", font=ctk.CTkFont(size=13)).pack()
        self.lang_var = ctk.StringVar(value="both")
        lang_row = ctk.CTkFrame(left, fg_color="transparent")
        lang_row.pack(pady=5)
        for text, val in [("English + Urdu", "both"), ("English", "en"), ("Urdu", "ur")]:
            ctk.CTkRadioButton(
                lang_row, text=text, variable=self.lang_var, value=val,
                font=ctk.CTkFont(size=12),
            ).pack(side="left", padx=8)

        ctk.CTkButton(
            left,
            text="Convert to Text",
            height=44,
            font=ctk.CTkFont(size=15, weight="bold"),
            fg_color="#2ecc71",
            hover_color="#27ae60",
            text_color="black",
            command=self._convert,
        ).pack(padx=14, pady=(12, 4), fill="x")

        # Divider
        ctk.CTkFrame(left, height=1, fg_color="gray30").pack(fill="x", padx=14, pady=10)

        # Gemini API key
        ctk.CTkLabel(
            left, text="Gemini API Key (Google AI Studio)",
            font=ctk.CTkFont(size=12, weight="bold"),
        ).pack()

        key_row = ctk.CTkFrame(left, fg_color="transparent")
        key_row.pack(padx=14, pady=5, fill="x")

        self.api_key_entry = ctk.CTkEntry(
            key_row,
            placeholder_text="Paste Gemini API key here...",
            show="*",
            font=ctk.CTkFont(size=11),
            height=34,
        )
        self.api_key_entry.pack(side="left", fill="x", expand=True, padx=(0, 6))

        ctk.CTkButton(
            key_row, text="👁", width=34, height=34,
            font=ctk.CTkFont(size=14), fg_color="gray30",
            hover_color="gray40", command=self._toggle_key_visibility,
        ).pack(side="left")

        ctk.CTkButton(
            left,
            text="⚖  Compare with Gemini AI",
            height=42,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color="#8e44ad",
            hover_color="#7d3c98",
            command=self._compare,
        ).pack(padx=14, pady=(6, 2), fill="x")

        self.convert_status = ctk.CTkLabel(
            left, text="", font=ctk.CTkFont(size=11), text_color="gray"
        )
        self.convert_status.pack()

        # ── Right panel (output area) ────────────────────────────────────────
        self.right = ctk.CTkFrame(body)
        self.right.pack(side="right", fill="both", expand=True, padx=(10, 0))

        self._build_single_output()

    def _build_single_output(self):
        for w in self.right.winfo_children():
            w.destroy()
        self._compare_mode = False

        ctk.CTkLabel(
            self.right, text="Extracted Text",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(pady=(14, 8))

        self.textbox = ctk.CTkTextbox(
            self.right, font=ctk.CTkFont(family="Segoe UI", size=13), wrap="word",
        )
        self.textbox.pack(fill="both", expand=True, padx=14)

        btn_row = ctk.CTkFrame(self.right, fg_color="transparent")
        btn_row.pack(pady=12)

        ctk.CTkButton(
            btn_row, text="Copy Text", width=130, height=38,
            font=ctk.CTkFont(size=13), command=self._copy_text,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_row, text="Save as Word (.docx)", width=170, height=38,
            font=ctk.CTkFont(size=13), fg_color="#1a5276",
            hover_color="#154360", command=self._save_word,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_row, text="Clear", width=90, height=38,
            font=ctk.CTkFont(size=13), fg_color="gray35",
            hover_color="gray25", command=self._clear,
        ).pack(side="left", padx=6)

    def _build_compare_output(self):
        for w in self.right.winfo_children():
            w.destroy()
        self._compare_mode = True

        # Column headers
        hdr = ctk.CTkFrame(self.right, fg_color="transparent")
        hdr.pack(fill="x", padx=14, pady=(14, 4))

        ctk.CTkLabel(
            hdr, text="EasyOCR  (Local / Offline)",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#2ecc71",
        ).pack(side="left", expand=True)

        ctk.CTkLabel(
            hdr, text="Gemini AI  (API)",
            font=ctk.CTkFont(size=13, weight="bold"), text_color="#3498db",
        ).pack(side="right", expand=True)

        # Two textboxes
        text_row = ctk.CTkFrame(self.right, fg_color="transparent")
        text_row.pack(fill="both", expand=True, padx=14)

        self.textbox = ctk.CTkTextbox(
            text_row, font=ctk.CTkFont(family="Segoe UI", size=12), wrap="word",
        )
        self.textbox.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.textbox_gv = ctk.CTkTextbox(
            text_row, font=ctk.CTkFont(family="Segoe UI", size=12), wrap="word",
        )
        self.textbox_gv.pack(side="right", fill="both", expand=True, padx=(5, 0))

        # Stats row
        self.stats_label = ctk.CTkLabel(
            self.right, text="",
            font=ctk.CTkFont(size=12, weight="bold"),
        )
        self.stats_label.pack(pady=(6, 2))

        btn_row = ctk.CTkFrame(self.right, fg_color="transparent")
        btn_row.pack(pady=8)

        ctk.CTkButton(
            btn_row, text="← Back to Single View", width=160, height=36,
            font=ctk.CTkFont(size=12), fg_color="gray35",
            hover_color="gray25", command=self._build_single_output,
        ).pack(side="left", padx=6)

        ctk.CTkButton(
            btn_row, text="Clear", width=90, height=36,
            font=ctk.CTkFont(size=12), fg_color="gray35",
            hover_color="gray25", command=self._clear,
        ).pack(side="left", padx=6)

    # ── Helpers ─────────────────────────────────────────────────────────────

    def _toggle_key_visibility(self):
        self.api_key_entry.configure(
            show="" if self.api_key_entry.cget("show") == "*" else "*"
        )

    # ── OCR model loader ────────────────────────────────────────────────────

    def _load_model_async(self):
        threading.Thread(target=self._load_model, daemon=True).start()

    def _load_model(self):
        try:
            self.reader = easyocr.Reader(["en", "ur"], gpu=False)
            self.after(0, lambda: self.status_badge.configure(
                text="✅  Ready", text_color="#2ecc71"
            ))
        except Exception as exc:
            self.after(0, lambda: self.status_badge.configure(
                text=f"❌  Model error: {exc}", text_color="red"
            ))

    # ── Image upload ────────────────────────────────────────────────────────

    def _pick_image(self):
        path = filedialog.askopenfilename(
            title="Select Paper Image",
            filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.tiff *.tif"), ("All", "*.*")],
        )
        if path:
            self.current_image_path = path
            self._show_preview(path)
            self.convert_status.configure(text="Image loaded ✔", text_color="#2ecc71")

    def _show_preview(self, path):
        try:
            img = Image.open(path)
            img.thumbnail((350, 220))
            photo = ImageTk.PhotoImage(img)
            for w in self.drop_zone.winfo_children():
                w.destroy()
            lbl = ctk.CTkLabel(self.drop_zone, image=photo, text="")
            lbl.image = photo
            lbl.pack(expand=True)
        except Exception as exc:
            self.convert_status.configure(text=f"Preview error: {exc}", text_color="red")

    # ── EasyOCR single convert ───────────────────────────────────────────────

    def _convert(self):
        if not self.current_image_path:
            messagebox.showwarning("No Image", "Please select an image first.")
            return
        if not self.reader:
            messagebox.showwarning("Please Wait", "OCR model is still loading.")
            return
        if self._compare_mode:
            self._build_single_output()
        self.convert_status.configure(text="⏳  Processing...", text_color="orange")
        threading.Thread(target=self._run_ocr, daemon=True).start()

    def _run_ocr(self):
        try:
            lang = self.lang_var.get()
            reader = (
                easyocr.Reader(["en"], gpu=False) if lang == "en"
                else easyocr.Reader(["ur"], gpu=False) if lang == "ur"
                else self.reader
            )
            results = reader.readtext(self.current_image_path)
            text = "\n".join(r[1] for r in results)
            self.after(0, lambda: self._show_result(text))
        except Exception as exc:
            self.after(0, lambda: self.convert_status.configure(
                text=f"Error: {exc}", text_color="red"
            ))

    def _show_result(self, text):
        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", text)
        self.convert_status.configure(text="✅  Done!", text_color="#2ecc71")

    # ── Comparison ───────────────────────────────────────────────────────────

    def _compare(self):
        if not self.current_image_path:
            messagebox.showwarning("No Image", "Please select an image first.")
            return
        if not self.reader:
            messagebox.showwarning("Please Wait", "OCR model is still loading.")
            return
        api_key = self.api_key_entry.get().strip()
        if not api_key:
            messagebox.showwarning("API Key Missing", "Please enter your Google Vision API key.")
            return

        self._build_compare_output()
        self.convert_status.configure(text="⏳  Running both engines...", text_color="orange")
        threading.Thread(target=self._run_compare, args=(api_key,), daemon=True).start()

    def _run_compare(self, api_key):
        results = {}

        # EasyOCR
        try:
            t0 = time.time()
            lang = self.lang_var.get()
            reader = (
                easyocr.Reader(["en"], gpu=False) if lang == "en"
                else easyocr.Reader(["ur"], gpu=False) if lang == "ur"
                else self.reader
            )
            ocr_out = reader.readtext(self.current_image_path)
            results["easy"] = "\n".join(r[1] for r in ocr_out)
            results["easy_time"] = round(time.time() - t0, 2)
        except Exception as exc:
            results["easy"] = f"EasyOCR error: {exc}"
            results["easy_time"] = 0

        # Gemini
        try:
            import mimetypes
            t0 = time.time()
            mime_type, _ = mimetypes.guess_type(self.current_image_path)
            if mime_type not in ("image/jpeg", "image/png", "image/gif", "image/webp"):
                mime_type = "image/jpeg"

            with open(self.current_image_path, "rb") as f:
                encoded = base64.b64encode(f.read()).decode("utf-8")

            url = (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                f"gemini-2.0-flash:generateContent?key={api_key}"
            )
            payload = {
                "contents": [{
                    "parts": [
                        {"inline_data": {"mime_type": mime_type, "data": encoded}},
                        {"text": (
                            "Extract all text from this image exactly as it appears. "
                            "Preserve line breaks and formatting. "
                            "Return only the extracted text, no commentary."
                        )},
                    ]
                }]
            }
            resp = requests.post(url, json=payload, timeout=30)
            data = resp.json()

            if "error" in data:
                results["gv"] = f"API Error: {data['error'].get('message', 'Unknown')}"
            else:
                parts = data["candidates"][0]["content"]["parts"]
                results["gv"] = parts[0].get("text", "(No text detected)")
            results["gv_time"] = round(time.time() - t0, 2)
        except Exception as exc:
            results["gv"] = f"Gemini error: {exc}"
            results["gv_time"] = 0

        self.after(0, lambda: self._show_compare_result(results))

    def _show_compare_result(self, results):
        easy_text = results.get("easy", "")
        gv_text = results.get("gv", "")
        easy_time = results.get("easy_time", 0)
        gv_time = results.get("gv_time", 0)

        self.textbox.delete("1.0", "end")
        self.textbox.insert("1.0", easy_text)

        self.textbox_gv.delete("1.0", "end")
        self.textbox_gv.insert("1.0", gv_text)

        sim = _similarity(easy_text, gv_text)
        color = "#2ecc71" if sim >= 70 else ("orange" if sim >= 40 else "#e74c3c")
        self.stats_label.configure(
            text=(
                f"Text Similarity: {sim}%   |   "
                f"EasyOCR: {easy_time}s   |   "
                f"Gemini: {gv_time}s"
            ),
            text_color=color,
        )
        self.convert_status.configure(text="✅  Comparison complete!", text_color="#2ecc71")

    # ── Output actions ───────────────────────────────────────────────────────

    def _copy_text(self):
        text = self.textbox.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("Empty", "No text to copy.")
            return
        self.clipboard_clear()
        self.clipboard_append(text)
        self.convert_status.configure(text="Copied to clipboard!", text_color="#2ecc71")

    def _save_word(self):
        text = self.textbox.get("1.0", "end-1c").strip()
        if not text:
            messagebox.showinfo("Empty", "No text to save.")
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word Document", "*.docx")],
            title="Save Document",
        )
        if path:
            try:
                doc = Document()
                doc.add_paragraph(text)
                doc.save(path)
                self.convert_status.configure(text="Saved successfully!", text_color="#2ecc71")
                messagebox.showinfo("Saved", f"File saved:\n{path}")
            except Exception as exc:
                messagebox.showerror("Error", str(exc))

    def _clear(self):
        self.current_image_path = None
        for w in self.drop_zone.winfo_children():
            w.destroy()
        self.img_hint = ctk.CTkLabel(
            self.drop_zone,
            text="📄\n\nClick  'Select Image'  below\nto upload a scanned paper",
            font=ctk.CTkFont(size=13),
            text_color="gray55",
            justify="center",
        )
        self.img_hint.pack(expand=True)

        self.textbox.delete("1.0", "end")
        if self._compare_mode and hasattr(self, "textbox_gv"):
            self.textbox_gv.delete("1.0", "end")
            self.stats_label.configure(text="")
        self.convert_status.configure(text="Cleared", text_color="gray")


if __name__ == "__main__":
    app = ComposerApp()
    app.mainloop()
