import os, json, threading, subprocess
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Progressbar, Combobox
from gtts import gTTS

# ================= EDGE TTS SAFE =================
def edge_tts(text, out, voice):
    if not text.strip():
        return False
    cmd = [
        "edge-tts",
        "--voice", voice,
        "--text", text,
        "--write-media", out
    ]
    try:
        subprocess.run(cmd, check=True,
                       stdout=subprocess.DEVNULL,
                       stderr=subprocess.DEVNULL)
        return True
    except:
        return False

# ================= APP =================
class LessonGeneratorPRO:
    def __init__(self, root):
        self.root = root
        root.title("EXCEL → LESSON (PRO STABLE)")
        root.geometry("760x520")
        self.ui()

    def ui(self):
        self.excel = self.row("Excel file:", self.pick_excel)
        self.out = self.row("Output folder:", self.pick_out)

        box = tk.LabelFrame(self.root, text="Audio Options", padx=10, pady=10)
        box.pack(fill=tk.X, pady=10)

        tk.Label(box, text="Generate audio (English):").grid(row=0, column=0, sticky="w")
        tk.Label(box, text="Will generate both: gTTS (Female US) and Edge TTS (Male US)").grid(row=0, column=1, sticky="w")

        self.progress = Progressbar(self.root, length=650)
        self.progress.pack(pady=15)

        tk.Button(self.root, text="🚀 CREATE LESSON",
                  bg="#0066CC", fg="white",
                  font=("Arial", 11, "bold"),
                  command=self.run).pack(pady=10)

    def row(self, label, cmd):
        f = tk.Frame(self.root)
        f.pack()
        tk.Label(f, text=label).pack(side=tk.LEFT)
        e = tk.Entry(f, width=65)
        e.pack(side=tk.LEFT, padx=5)
        tk.Button(f, text="Browse", command=cmd).pack(side=tk.LEFT)
        return e

    def pick_excel(self):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if f:
            self.excel.delete(0, tk.END)
            self.excel.insert(0, f)

    def pick_out(self):
        d = filedialog.askdirectory()
        if d:
            self.out.delete(0, tk.END)
            self.out.insert(0, d)

    # ================= CORE =================
    def run(self):
        threading.Thread(target=self.build, daemon=True).start()

    def set_progress_value(self, value):
        try:
            self.root.after(0, lambda: self.progress.__setitem__('value', value))
        except Exception:
            pass

    def build(self):
        try:
            if not self.excel.get().strip():
                self.root.after(0, lambda: messagebox.showwarning("Missing file", "Please choose an Excel file"))
                return
            if not self.out.get().strip():
                self.root.after(0, lambda: messagebox.showwarning("Missing output", "Please choose an output folder"))
                return

            df = pd.read_excel(self.excel.get()).fillna("")
            lesson_name = os.path.splitext(os.path.basename(self.excel.get()))[0]

            # Use the folder the user selected directly (no extra LESSON folder)
            base = self.out.get()
            os.makedirs(base, exist_ok=True)
            audio_dir = os.path.join(base, f"audio_{lesson_name}")
            os.makedirs(audio_dir, exist_ok=True)

            lesson_data = []
            mapping = []

            # compute total tasks accurately (generate two engines per English text)
            total = 0
            for _, r in df.iterrows():
                row_list = r.tolist()
                while len(row_list) < 6:
                    row_list.append("")
                row = [str(x) for x in row_list[:6]]
                # for each English item (front or example) we will generate 2 files (gTTS + Edge)
                if row[0].strip():
                    total += 2
                if row[3].strip():
                    total += 2

            if total == 0:
                self.root.after(0, lambda: messagebox.showinfo("No tasks", "No audio to generate (empty inputs or options)."))
                return

            done = 0

            for i, r in df.iterrows():
                row_list = r.tolist()
                while len(row_list) < 6:
                    row_list.append("")
                row = [str(x) for x in row_list[:6]]
                lesson_data.append(row)

                # ENGLISH: generate both gTTS (female) and Edge TTS (male) for each text field
                for text, typ in [(row[0], "en"), (row[3], "ex_en")]:
                    if not str(text).strip():
                        continue
                    # gTTS file
                    g_file = f"gtts_{typ}_{i}.mp3"
                    g_path = os.path.join(audio_dir, g_file)
                    try:
                        gTTS(text, lang="en").save(g_path)
                        mapping.append({
                            "id": f"{typ}_{i}",
                            "text": text,
                            "file": f"{os.path.basename(audio_dir)}/{g_file}",
                            "lang": "en",
                            "type": typ,
                            "engine": "gTTS",
                            "voice": "Female US (gTTS)"
                        })
                        done += 1
                        self.set_progress_value(done / total * 100)
                    except Exception:
                        pass

                    # Edge TTS file
                    e_file = f"edge_{typ}_{i}.mp3"
                    e_path = os.path.join(audio_dir, e_file)
                    try:
                        ok_edge = edge_tts(text, e_path, "en-US-GuyNeural")
                        if ok_edge:
                            mapping.append({
                                "id": f"{typ}_{i}",
                                "text": text,
                                "file": f"{os.path.basename(audio_dir)}/{e_file}",
                                "lang": "en",
                                "type": typ,
                                "engine": "Edge",
                                "voice": "en-US-GuyNeural"
                            })
                        done += 1
                        self.set_progress_value(done / total * 100)
                    except Exception:
                        done += 1
                        self.set_progress_value(done / total * 100)

            # ===== SAVE JSON =====
            lesson_file = os.path.join(base, f"{lesson_name}.json")
            with open(lesson_file, "w", encoding="utf-8") as f:
                json.dump(lesson_data, f, ensure_ascii=False, indent=2)

            mapping_file = os.path.join(base, f"mapping_{lesson_name}.json")
            with open(mapping_file, "w", encoding="utf-8") as f:
                json.dump(mapping, f, ensure_ascii=False, indent=2)

            self.root.after(0, lambda: messagebox.showinfo(
                "SUCCESS",
                f"✔ Lesson created successfully!\n\n"
                f"- {lesson_file}\n"
                f"- {mapping_file}\n"
                f"- {os.path.basename(audio_dir)}/"
            ))

        finally:
            self.set_progress_value(0)

# ================= RUN =================
if __name__ == "__main__":
    root = tk.Tk()
    LessonGeneratorPRO(root)
    root.mainloop()
