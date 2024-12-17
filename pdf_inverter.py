import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import sys
import ctypes

class PDFToolApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Merger & Reverser")
        self.root.geometry("700x600")
        self.root.resizable(True, True)  # Allow window to resize

        # Color Palette
        self.setup_colors()

        # PDF Lists
        self.merge_pdf_list = []

        # Setup Styles
        self.setup_styles()

        # Create UI
        self.create_ui()

    def minimize_console(self):
        if sys.platform == "win32":
            try:
                kernel32 = ctypes.WinDLL('kernel32')
                user32 = ctypes.WinDLL('user32')
                hwnd = kernel32.GetConsoleWindow()

                if hwnd != 0:
                    user32.ShowWindow(hwnd, 2)  # Minimize the console window
            except AttributeError:
                print("GetConsoleWindow function not found. The environment may not support this.")
            except Exception as e:
                print(f"Error minimizing console: {e}")
        else:
            print("Minimizing console is only supported on Windows.")
        
    def setup_colors(self):
        """Define color palette"""
        self.BG_COLOR = "#d3d3d3"
        self.FONT_COLOR = "#000000"  # Black
        self.INPUT_FOCUS = "#2d8cf0"
        self.BUTTON_BG = "#323232"
        self.WHITE = "#ffffff"

    def setup_styles(self):
        """Set text style globally to bold and black"""
        self.bold_font = ("Segoe UI", 11, "bold")
        self.title_font = ("Segoe UI", 20, "bold")

        self.style = ttk.Style()
        self.style.configure(
            "TButton",
            font=self.bold_font,
            background=self.BUTTON_BG,
            foreground=self.FONT_COLOR,  # Set button text color to black
            padding=10,
        )

    def create_ui(self):
        """Create the main UI"""
        # Main container frame
        self.main_frame = tk.Frame(self.root, bg=self.BG_COLOR, padx=20, pady=20)
        self.main_frame.pack(expand=True, fill=tk.BOTH)

        # Title
        title_label = tk.Label(
            self.main_frame,
            text="PDF Merger & Reverser",
            font=self.title_font,
            fg=self.FONT_COLOR,
            bg=self.BG_COLOR,
        )
        title_label.pack(pady=10)

        # Separator
        tk.Frame(self.main_frame, bg=self.FONT_COLOR, height=2).pack(fill=tk.X, pady=10)

        # Add PDFs Section
        add_frame = tk.Frame(self.main_frame, bg=self.BG_COLOR)
        add_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            add_frame, text="Merge PDFs:", bg=self.BG_COLOR, fg=self.FONT_COLOR, font=self.bold_font
        ).pack(anchor="w")

        self.pdf_listbox = tk.Listbox(
            add_frame,
            bg=self.WHITE,
            fg=self.FONT_COLOR,
            font=self.bold_font,
            selectbackground=self.INPUT_FOCUS,
            relief=tk.FLAT,
            height=5,
        )
        self.pdf_listbox.pack(fill=tk.X, pady=5)

        button_frame = tk.Frame(add_frame, bg=self.BG_COLOR)
        button_frame.pack(fill=tk.X, expand=True, pady=5)

        # Create buttons and pack them with fill and side=LEFT to ensure proportional sizing
        ttk.Button(button_frame, text="Add PDFs", command=self.add_multiple_pdfs).pack(side=tk.LEFT, fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="Remove", command=self.remove_pdf).pack(side=tk.LEFT, fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="Move Up", command=self.move_up).pack(side=tk.LEFT, fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="Move Down", command=self.move_down).pack(side=tk.LEFT, fill=tk.X, padx=5, pady=5)
        ttk.Button(button_frame, text="Merge", command=self.merge_pdfs).pack(side=tk.RIGHT, padx=5, pady=5)

        # Separator
        tk.Frame(self.main_frame, bg=self.FONT_COLOR, height=2).pack(fill=tk.X, pady=10)

        # Reverse PDFs Section
        reverse_frame = tk.Frame(self.main_frame, bg=self.BG_COLOR)
        reverse_frame.pack(fill=tk.X, pady=10)

        tk.Label(
            reverse_frame, text="Reverse PDF:", bg=self.BG_COLOR, fg=self.FONT_COLOR, font=self.bold_font
        ).pack(anchor="w")

        self.reverse_input = tk.Entry(
            reverse_frame,
            font=self.bold_font,
            fg=self.FONT_COLOR,
            bg=self.WHITE,
            relief=tk.FLAT,
            highlightthickness=2,
            highlightbackground=self.FONT_COLOR,
            highlightcolor=self.INPUT_FOCUS,
        )
        self.reverse_input.pack(fill=tk.X, pady=5)

        ttk.Button(reverse_frame, text="Browse", command=self.select_reverse_pdf).pack(side=tk.LEFT, padx=5, pady=5)
        ttk.Button(reverse_frame, text="Reverse", command=self.reverse_pdf).pack(side=tk.RIGHT, padx=5, pady=5)

        # Progress Bar
        self.progress_bar = ttk.Progressbar(self.main_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=20)

    def add_multiple_pdfs(self):
        file_paths = filedialog.askopenfilenames(title="Select PDFs to Merge", filetypes=[("PDF files", "*.pdf")])
        for file_path in file_paths:
            self.merge_pdf_list.append(file_path)
            self.pdf_listbox.insert(tk.END, os.path.basename(file_path))

    def remove_pdf(self):
        try:
            selected_index = self.pdf_listbox.curselection()[0]
            self.pdf_listbox.delete(selected_index)
            del self.merge_pdf_list[selected_index]
        except IndexError:
            messagebox.showwarning("Warning", "Please select a PDF to remove.")

    def move_up(self):
        try:
            selected_index = self.pdf_listbox.curselection()[0]
            if selected_index > 0:
                # Swap in the list
                self.merge_pdf_list[selected_index], self.merge_pdf_list[selected_index - 1] = \
                    self.merge_pdf_list[selected_index - 1], self.merge_pdf_list[selected_index]
                # Update the listbox display
                self.update_pdf_listbox()
                self.pdf_listbox.select_set(selected_index - 1)
        except IndexError:
            messagebox.showwarning("Warning", "Please select a PDF to move.")

    def move_down(self):
        try:
            selected_index = self.pdf_listbox.curselection()[0]
            if selected_index < len(self.merge_pdf_list) - 1:
                # Swap in the list
                self.merge_pdf_list[selected_index], self.merge_pdf_list[selected_index + 1] = \
                    self.merge_pdf_list[selected_index + 1], self.merge_pdf_list[selected_index]
                # Update the listbox display
                self.update_pdf_listbox()
                self.pdf_listbox.select_set(selected_index + 1)
        except IndexError:
            messagebox.showwarning("Warning", "Please select a PDF to move.")

    def update_pdf_listbox(self):
        """Update the listbox to reflect the new order of PDFs"""
        self.pdf_listbox.delete(0, tk.END)
        for file in self.merge_pdf_list:
            self.pdf_listbox.insert(tk.END, os.path.basename(file))

    def merge_pdfs(self):
        if not self.merge_pdf_list:
            messagebox.showwarning("Warning", "No PDFs selected to merge.")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Merged PDF As"
        )
        if output_path:
            merger = PdfMerger()
            for pdf_path in self.merge_pdf_list:
                merger.append(pdf_path)
            with open(output_path, "wb") as output_file:
                merger.write(output_file)

            # Remove the selected files after merging
            # for pdf_path in self.merge_pdf_list:
            #     os.remove(pdf_path)
            
            self.merge_pdf_list.clear()  # Clear the list of selected PDFs
            self.update_pdf_listbox()  # Update the listbox
            messagebox.showinfo("Success", f"PDFs merged successfully!\nSaved to: {output_path}")

    def select_reverse_pdf(self):
        file_path = filedialog.askopenfilename(title="Select PDF to Reverse", filetypes=[("PDF files", "*.pdf")])
        if file_path:
            self.reverse_input.delete(0, tk.END)
            self.reverse_input.insert(0, file_path)

    def reverse_pdf(self):
        input_path = self.reverse_input.get()
        if not input_path or not os.path.exists(input_path):
            messagebox.showerror("Error", "Please select a valid PDF file!")
            return

        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], title="Save Reversed PDF As"
        )
        if output_path:
            reader = PdfReader(input_path)
            writer = PdfWriter()

            # Reverse the pages
            for page_num in range(len(reader.pages) - 1, -1, -1):
                writer.add_page(reader.pages[page_num])

            with open(output_path, "wb") as output_file:
                writer.write(output_file)

            # Remove the original file after reversing
            # os.remove(input_path)

            messagebox.showinfo("Success", f"PDF reversed successfully!\nSaved to: {output_path}")


def main():
    root = tk.Tk()
    app = PDFToolApp(root)
    # Minimize the console window after creating the Tkinter window
    app.minimize_console()
    root.mainloop()


if __name__ == "__main__":
    main()