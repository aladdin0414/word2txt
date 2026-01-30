#!/usr/bin/env python3
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
import sys

# Import extraction logic from word2txt.py
try:
    from word2txt import extract_docx_text
except ImportError:
    # Fallback if word2txt.py is not in path or has import issues
    sys.path.append(str(Path(__file__).resolve().parent))
    from word2txt import extract_docx_text


class WordTxtMergeApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Word to Text Merge Utility")
        self.root.geometry("600x450")

        # Variables
        self.selected_dir = tk.StringVar()
        self.files_found = []

        # UI Components
        self.create_widgets()

        # Set default directory to ./word if it exists
        base_dir = Path(__file__).resolve().parent
        default_word_dir = base_dir / "word"
        if default_word_dir.exists() and default_word_dir.is_dir():
            self.selected_dir.set(str(default_word_dir))
            # Pre-load files in default dir implies "Directory Mode", but user wants "File Mode".
            # We will just list them as if selected.
            self.scan_files(sorted(list(default_word_dir.glob("*.docx"))))

    def create_widgets(self):
        # 1. Directory Selection Area
        frame_top = tk.Frame(self.root, pady=10)
        frame_top.pack(fill=tk.X, padx=10)

        btn_select = tk.Button(frame_top, text="Select Files", command=self.select_files)
        btn_select.pack(side=tk.LEFT, padx=(0, 5))

        btn_remove = tk.Button(frame_top, text="Remove Selected", command=self.remove_selected)
        btn_remove.pack(side=tk.LEFT, padx=(0, 5))

        btn_clear = tk.Button(frame_top, text="Clear All", command=self.clear_all)
        btn_clear.pack(side=tk.LEFT, padx=(0, 10))

        lbl_path = tk.Label(frame_top, textvariable=self.selected_dir, fg="blue", wraplength=300, justify=tk.LEFT)
        lbl_path.pack(side=tk.LEFT, fill=tk.X)

        # 2. File List Area
        frame_list = tk.LabelFrame(self.root, text="Files Selected", padx=5, pady=5)
        frame_list.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.list_files = tk.Listbox(frame_list, selectmode=tk.EXTENDED)
        scrollbar = tk.Scrollbar(frame_list, orient=tk.VERTICAL, command=self.list_files.yview)
        
        self.list_files.config(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.list_files.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # 3. Action Area
        frame_bottom = tk.Frame(self.root, pady=10)
        frame_bottom.pack(fill=tk.X, padx=10)

        self.btn_run = tk.Button(frame_bottom, text="Extract & Merge", command=self.start_processing, state=tk.DISABLED, bg="#dddddd")
        self.btn_run.pack(side=tk.RIGHT)

        self.status_label = tk.Label(frame_bottom, text="Ready", fg="gray")
        self.status_label.pack(side=tk.LEFT)

    def select_files(self):
        initial = self.selected_dir.get()
        if not initial:
            base_dir = Path(__file__).resolve().parent
            initial = str(base_dir / "word")

        # Handle case where initial is a directory or a file path string from previous selection
        initial_dir = initial
        if not Path(initial_dir).is_dir():
             initial_dir = str(Path(initial_dir).parent)

        files = filedialog.askopenfilenames(
            initialdir=initial_dir,
            title="Select Word Files",
            filetypes=(("Word files", "*.docx"), ("All files", "*.*"))
        )
        if files:
            # Append new files to existing ones or replace? 
            # User might want to accumulate. "Select Files" usually adds.
            # But previously we replaced. Let's make it ADD uniquely.
            new_files = [Path(f) for f in files]
            
            # Combine unique files
            current_paths = {p.resolve() for p in self.files_found}
            for p in new_files:
                if p.resolve() not in current_paths:
                    self.files_found.append(p)
            
            self.files_found = sorted(self.files_found, key=lambda p: p.name)
            
            if self.files_found:
                 self.selected_dir.set(str(self.files_found[0].parent)) 
            
            self.refresh_list()

    def remove_selected(self):
        selected_indices = self.list_files.curselection()
        if not selected_indices:
            return

        # Indices are returned as strings in some tk versions, ensure int
        indices = [int(i) for i in selected_indices]
        
        # Remove from files_found (need to map listbox index to list index)
        # Since listbox matches files_found order:
        files_to_remove = [self.files_found[i] for i in indices]
        
        for f in files_to_remove:
            self.files_found.remove(f)
            
        self.refresh_list()

    def clear_all(self):
        self.files_found = []
        self.refresh_list()

    def refresh_list(self):
        self.list_files.delete(0, tk.END)
        for p in self.files_found:
            self.list_files.insert(tk.END, p.name)
            
        if self.files_found:
            self.btn_run.config(state=tk.NORMAL)
            self.status_label.config(text=f"Selected {len(self.files_found)} files.")
        else:
            self.btn_run.config(state=tk.DISABLED)
            self.status_label.config(text="No files selected.")

    def scan_files(self, files: list[Path]):
        # This implementation is now redundant or can be just an alias to populate files_found and refresh
        self.files_found = sorted(files, key=lambda p: p.name)
        self.refresh_list()

    def start_processing(self):
        self.btn_run.config(state=tk.DISABLED)
        self.status_label.config(text="Processing...")
        
        # Run in a separate thread to keep UI responsive
        thread = threading.Thread(target=self.process_files)
        thread.start()

    def process_files(self):
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_dir = Path(__file__).resolve().parent
            merge_dir = base_dir / "merge"
            merge_dir.mkdir(exist_ok=True, parents=True)
            
            output_file = merge_dir / f"merged_{timestamp}.txt"
            
            merged_content = []
            
            success_count = 0
            
            for i, file_path in enumerate(self.files_found):
                # Update status
                self.update_status(f"Processing {i+1}/{len(self.files_found)}: {file_path.name}")
                
                text_obj = extract_docx_text(file_path)
                if text_obj:
                    content = text_obj.to_text()
                    merged_content.append(content)
                    success_count += 1
                
                # Add separator
                if i < len(self.files_found) - 1:
                    merged_content.append("\n\n" + "-"*40 + "\n\n")

            # Write to file
            self.update_status("Saving merged file...")
            output_file.write_text("".join(merged_content), encoding="utf-8")
            
            self.root.after(0, lambda: self.finish_processing(success_count, output_file))
            
        except Exception as e:
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.root.after(0, lambda: self.reset_ui())

    def update_status(self, text):
        self.root.after(0, lambda: self.status_label.config(text=text))

    def finish_processing(self, count, out_path):
        messagebox.showinfo("Success", f"Merged {count} files.\nSaved to:\n{out_path}")
        self.status_label.config(text=f"Done. Saved to {out_path.name}")
        self.btn_run.config(state=tk.NORMAL)

    def reset_ui(self):
        self.btn_run.config(state=tk.NORMAL)
        self.status_label.config(text="Ready")


def main():
    root = tk.Tk()
    app = WordTxtMergeApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
