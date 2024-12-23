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
        self.root.geometry("800x650")  # Increased width to 800px
        self.root.resizable(False, False)  # Disable window resizing

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
        self.PROGRESS_COLOR = "#4CAF50"  # Green color for progress bar

    def setup_styles(self):
        """Set text style globally to bold and black"""
        self.bold_font = ("Segoe UI", 11, "bold")
        self.title_font = ("Segoe UI", 20, "bold")

        self.style = ttk.Style()
        
        # Configure button style
        self.style.configure(
            "TButton",
            font=self.bold_font,
            background=self.BUTTON_BG,
            foreground=self.FONT_COLOR,
            padding=10,
        )
        
        # Configure progress bar style - Windows-like appearance
        self.style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor='#E0E0E0',     # Light gray background
            background='#06B025',       # Windows-like green color
            darkcolor='#06B025',        # Consistent color when filled
            lightcolor='#06B025',       # Consistent color when filled
            bordercolor='#ACACAC',      # Border color
            thickness=25                # Increased height
        )

    def create_ui(self):
        """Create the main UI"""
        # Main container frame
        self.main_frame = tk.Frame(self.root, bg=self.BG_COLOR, padx=30, pady=20)
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

        # First row of buttons
        button_frame_1 = tk.Frame(add_frame, bg=self.BG_COLOR)
        button_frame_1.pack(fill=tk.X, expand=True, pady=5)

        # Create first row of buttons with equal spacing
        buttons_1 = [
            ("Add PDFs", self.add_multiple_pdfs),
            ("Remove", self.remove_pdf),
            ("Move Up", self.move_up),
            ("Move Down", self.move_down),
            ("Reverse List", self.reverse_list)
        ]

        for text, command in buttons_1:
            ttk.Button(button_frame_1, text=text, command=command).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Second row of buttons
        button_frame_2 = tk.Frame(add_frame, bg=self.BG_COLOR)
        button_frame_2.pack(fill=tk.X, expand=True, pady=10)  # Increased vertical padding

        # Add a frame for centering the bottom buttons
        center_frame = tk.Frame(button_frame_2, bg=self.BG_COLOR)
        center_frame.pack(expand=True)

        # Create second row of buttons centered
        ttk.Button(center_frame, text="Clear List", command=self.clear_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(center_frame, text="Merge", command=self.merge_pdfs).pack(side=tk.LEFT, padx=5)


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

        # Progress Bar Frame with Label
        progress_frame = tk.Frame(self.main_frame, bg=self.BG_COLOR)
        progress_frame.pack(fill=tk.X, pady=(20, 10))  # Added more top padding
        
       # Progress Bar Frame (removed label since we won't show percentage)
        progress_frame = tk.Frame(self.main_frame, bg=self.BG_COLOR)
        progress_frame.pack(fill=tk.X, pady=(20, 10))

        # Progress Bar
        self.progress_bar = ttk.Progressbar(
            progress_frame,
            style="Custom.Horizontal.TProgressbar",
            mode="determinate",
            length=740
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

    def reverse_list(self):
        """Reverse the order of PDFs in the list"""
        if not self.merge_pdf_list:
            messagebox.showwarning("Warning", "No PDFs in the list to reverse.")
            return
            
        self.merge_pdf_list.reverse()
        self.update_pdf_listbox()
        
    def clear_list(self):
        """Clear all PDFs from the list"""
        if not self.merge_pdf_list:
            messagebox.showwarning("Warning", "List is already empty.")
            return
            
        self.merge_pdf_list.clear()
        self.pdf_listbox.delete(0, tk.END)

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
            try:
                merger = PdfMerger()
                total_files = len(self.merge_pdf_list)
                
                # Reset progress bar
                self.progress_bar["value"] = 0
                self.root.update_idletasks()
                
                # Disable all buttons during merging
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Button):
                        widget.configure(state='disabled')
                
                for i, pdf_path in enumerate(self.merge_pdf_list):
                    # Update progress before merging each file
                    progress = ((i) / total_files) * 100
                    self.progress_bar["value"] = progress
                    self.root.update_idletasks()
                    # self.root.after(200)  # Delay for visibility
                    
                    # Merge the PDF
                    merger.append(pdf_path)
                    
                    # Update progress after merging each file
                    progress = ((i + 1) / total_files) * 100
                    self.progress_bar["value"] = progress
                    self.root.update_idletasks()
                    # self.root.after(200)  # Delay for visibility

                # Show full progress before writing
                self.progress_bar["value"] = 100
                self.root.update_idletasks()
                # self.root.after(500)  # Longer delay at completion
                
                # Write the merged PDF
                with open(output_path, "wb") as output_file:
                    merger.write(output_file)
                
                # Clear and update UI
                self.merge_pdf_list.clear()
                self.update_pdf_listbox()
                
                # Re-enable all buttons
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Button):
                        widget.configure(state='normal')
                
                # Show success message
                messagebox.showinfo("Success", f"PDFs merged successfully!\nSaved to: {output_path}")
                
                # Reset progress bar after success
                self.root.after(1000, lambda: self.progress_bar.configure(value=0))
                
            except Exception as e:
                # Re-enable buttons in case of error
                for widget in self.root.winfo_children():
                    if isinstance(widget, ttk.Button):
                        widget.configure(state='normal')
                        
                messagebox.showerror("Error", f"An error occurred while merging PDFs:\n{str(e)}")
                self.progress_bar["value"] = 0
                self.root.update_idletasks()

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