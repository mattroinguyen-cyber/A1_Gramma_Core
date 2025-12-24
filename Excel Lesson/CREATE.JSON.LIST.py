import tkinter as tk
from tkinter import filedialog, messagebox
import json
import os

class LessonGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Lesson JSON Generator")
        self.geometry("650x500")
        
        # Danh sách chứa thông tin bài học
        self.lesson_list = []
        
        # Các biến lưu đường dẫn được chọn
        self.lesson_file_path = tk.StringVar()
        self.mapping_file_path = tk.StringVar()
        self.audio_folder_path = tk.StringVar()
        
        # Giao diện chọn file
        tk.Label(self, text="Lesson JSON File:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.lesson_file_path, width=50).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=self.browse_lesson_file).grid(row=0, column=2, padx=10, pady=5)
        
        tk.Label(self, text="Mapping JSON File:").grid(row=1, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.mapping_file_path, width=50).grid(row=1, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=self.browse_mapping_file).grid(row=1, column=2, padx=10, pady=5)
        
        tk.Label(self, text="Audio Folder:").grid(row=2, column=0, sticky="w", padx=10, pady=5)
        tk.Entry(self, textvariable=self.audio_folder_path, width=50).grid(row=2, column=1, padx=10, pady=5)
        tk.Button(self, text="Browse", command=self.browse_audio_folder).grid(row=2, column=2, padx=10, pady=5)
        
        # Nút thêm bài học vào danh sách
        tk.Button(self, text="Add Lesson", command=self.add_lesson, width=20).grid(row=3, column=1, pady=10)
        
        # Nút xóa bài học đã chọn
        tk.Button(self, text="Delete Lesson", command=self.delete_lesson, width=20).grid(row=4, column=1, pady=10)
        
        # Nút nhập thêm từ file JSON
        tk.Button(self, text="Load from JSON", command=self.load_from_json, width=20).grid(row=5, column=1, pady=10)
        
        # Listbox hiển thị các bài học đã thêm
        tk.Label(self, text="Added Lessons:").grid(row=6, column=0, sticky="w", padx=10, pady=5)
        self.lesson_listbox = tk.Listbox(self, width=80, height=10)
        self.lesson_listbox.grid(row=7, column=0, columnspan=3, padx=10, pady=5)
        
        # Nút xuất file JSON tổng hợp
        tk.Button(self, text="Generate lessons.json", command=self.generate_json, width=25).grid(row=8, column=1, pady=10)

    def browse_lesson_file(self):
        filename = filedialog.askopenfilename(
            title="Select Lesson JSON File", 
            filetypes=[("JSON files", "*.json")]
        )
        if filename:
            self.lesson_file_path.set(filename)
    
    def browse_mapping_file(self):
        filename = filedialog.askopenfilename(
            title="Select Mapping JSON File", 
            filetypes=[("JSON files", "*.json")]
        )
        if filename:
            self.mapping_file_path.set(filename)
    
    def browse_audio_folder(self):
        foldername = filedialog.askdirectory(title="Select Audio Folder")
        if foldername:
            self.audio_folder_path.set(foldername)
    
    def add_lesson(self):
        lesson_file = self.lesson_file_path.get()
        mapping_file = self.mapping_file_path.get()
        audio_folder = self.audio_folder_path.get()
        
        if not lesson_file or not mapping_file or not audio_folder:
            messagebox.showwarning("Missing data", "Please select all required files/folder.")
            return
        
        # Lấy tên file (basename) để lưu lại, vì khi deploy thường các file này nằm cùng folder với ứng dụng
        lesson_name = os.path.basename(lesson_file)
        mapping_name = os.path.basename(mapping_file)
        audio_name = os.path.basename(audio_folder)
        
        # Tạo một dictionary chứa thông tin bài học
        lesson_entry = {
            "name": lesson_name,
            "mapping": mapping_name,
            "audio": audio_name
        }
        self.lesson_list.append(lesson_entry)
        self.update_lesson_listbox()
        
        # Xóa các trường đã chọn để sẵn sàng cho bài học tiếp theo
        self.lesson_file_path.set("")
        self.mapping_file_path.set("")
        self.audio_folder_path.set("")
    
    def update_lesson_listbox(self):
        self.lesson_listbox.delete(0, tk.END)
        for lesson in self.lesson_list:
            display_text = f"Lesson: {lesson['name']}, Mapping: {lesson['mapping']}, Audio: {lesson['audio']}"
            self.lesson_listbox.insert(tk.END, display_text)

    def delete_lesson(self):
        selected_index = self.lesson_listbox.curselection()
        if not selected_index:
            messagebox.showwarning("No selection", "Please select a lesson to delete.")
            return
        
        selected_index = selected_index[0]
        del self.lesson_list[selected_index]
        self.update_lesson_listbox()

    def load_from_json(self):
        filepath = filedialog.askopenfilename(
            title="Open JSON File", 
            filetypes=[("JSON files", "*.json")]
        )
        if not filepath:
            return

        try:
            with open(filepath, "r", encoding="utf-8") as f:
                data = json.load(f)
            
            if "lessons" in data:
                for lesson in data["lessons"]:
                    self.lesson_list.append(lesson)
                self.update_lesson_listbox()
            else:
                messagebox.showwarning("Invalid File", "The selected file does not contain lessons data.")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load the file: {e}")

    def generate_json(self):
        if not self.lesson_list:
            messagebox.showwarning("No lessons", "No lessons have been added.")
            return
        
        final_data = {
            "lessons": self.lesson_list
        }
        
        save_path = filedialog.asksaveasfilename(
            title="Save lessons.json", 
            defaultextension=".json", 
            filetypes=[("JSON files", "*.json")]
        )
        if save_path:
            with open(save_path, "w", encoding="utf-8") as f:
                json.dump(final_data, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Success", f"Generated lessons.json at {save_path}")

if __name__ == "__main__":
    app = LessonGUI()
    app.mainloop()
