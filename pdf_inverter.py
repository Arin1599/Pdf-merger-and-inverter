import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
import sys
import ctypes
from PIL import Image, ImageTk
import fitz  # PyMuPDF for better PDF rendering
import io
import tempfile

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
        ttk.Button(center_frame, text="Edit Pages", command=self.open_page_editor).pack(side=tk.LEFT, padx=5)


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

    def open_page_editor(self):
        """Open the PDF page editor window"""
        if not self.merge_pdf_list:
            messagebox.showwarning("Warning", "Please add PDFs to the list first.")
            return
        
        editor = PDFPageEditor(self.root, self.merge_pdf_list.copy())


class PDFPageEditor:
    def __init__(self, parent, pdf_files):
        self.parent = parent
        self.pdf_files = pdf_files
        self.pages = []  # List of page objects with metadata
        self.selected_page = None
        self.drag_data = {"x": 0, "y": 0, "item": None, "widget": None}
        self.thumbnail_frames = []  # Store references to thumbnail frames
        
        # Create editor window
        self.window = tk.Toplevel(parent)
        self.window.title("PDF Page Editor")
        self.window.geometry("1200x800")
        self.window.resizable(True, True)
        
        # Load all pages from PDFs
        self.load_pages()
        
        # Create UI
        self.create_editor_ui()
        
        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()
        
    def load_pages(self):
        """Load all pages from the PDF files"""
        self.pages = []
        for pdf_path in self.pdf_files:
            try:
                doc = fitz.open(pdf_path)
                for page_num in range(len(doc)):
                    page_data = {
                        'pdf_path': pdf_path,
                        'page_num': page_num,
                        'type': 'pdf',
                        'doc': doc,
                        'thumbnail': None
                    }
                    self.pages.append(page_data)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load {pdf_path}: {str(e)}")
    
    def create_editor_ui(self):
        """Create the page editor interface"""
        # Main container
        main_frame = tk.Frame(self.window, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_label = tk.Label(main_frame, text="PDF Page Editor", 
                              font=("Segoe UI", 16, "bold"), bg="#f0f0f0")
        title_label.pack(pady=(0, 10))
        
        # Toolbar
        toolbar = tk.Frame(main_frame, bg="#f0f0f0")
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(toolbar, text="Add Image", command=self.add_image).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Add PDF", command=self.add_pdf).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Delete Selected", command=self.delete_selected).pack(side=tk.LEFT, padx=5)
        ttk.Button(toolbar, text="Save PDF", command=self.save_pdf).pack(side=tk.RIGHT, padx=5)
        ttk.Button(toolbar, text="Cancel", command=self.close_editor).pack(side=tk.RIGHT, padx=5)
        
        # Scrollable canvas for thumbnails
        canvas_frame = tk.Frame(main_frame, bg="#f0f0f0")
        canvas_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create canvas with scrollbars
        self.canvas = tk.Canvas(canvas_frame, bg="white", highlightthickness=1, 
                               highlightbackground="#cccccc")
        v_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(canvas_frame, orient=tk.HORIZONTAL, command=self.canvas.xview)
        
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack scrollbars and canvas
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        # Create scrollable frame inside canvas
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")
        self.canvas_window = self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        
        # Bind events
        self.scrollable_frame.bind("<Configure>", self.on_frame_configure)
        self.canvas.bind("<Configure>", self.on_canvas_configure)
        self.canvas.bind("<MouseWheel>", self.on_mousewheel)
        
        # Generate thumbnails and display pages
        self.refresh_thumbnails()
    
    def on_frame_configure(self, event):
        """Update scroll region when frame size changes"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
    
    def on_canvas_configure(self, event):
        """Update canvas window width when canvas is resized"""
        canvas_width = event.width
        self.canvas.itemconfig(self.canvas_window, width=canvas_width)
    
    def on_mousewheel(self, event):
        """Handle mouse wheel scrolling"""
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    def generate_thumbnail(self, page_data, size=(150, 200)):
        """Generate thumbnail for a page"""
        try:
            if page_data['type'] == 'pdf':
                # Get page from PDF
                page = page_data['doc'][page_data['page_num']]
                # Render page as image
                mat = fitz.Matrix(1.0, 1.0)  # Scale factor
                pix = page.get_pixmap(matrix=mat)
                img_data = pix.tobytes("ppm")
                img = Image.open(io.BytesIO(img_data))
            elif page_data['type'] == 'image':
                img = Image.open(page_data['image_path'])
            
            # Resize to thumbnail size while maintaining aspect ratio
            img.thumbnail(size, Image.Resampling.LANCZOS)
            
            # Create a white background and paste the image centered
            thumb = Image.new('RGB', size, 'white')
            x = (size[0] - img.width) // 2
            y = (size[1] - img.height) // 2
            thumb.paste(img, (x, y))
            
            return ImageTk.PhotoImage(thumb)
        except Exception as e:
            print(f"Error generating thumbnail: {e}")
            # Return a placeholder image
            placeholder = Image.new('RGB', size, '#f0f0f0')
            return ImageTk.PhotoImage(placeholder)
    
    def refresh_thumbnails(self):
        """Refresh the thumbnail display with drag-and-drop and insert buttons"""
        # Clear existing thumbnails
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        
        self.thumbnail_frames = []
        
        # Create grid layout: each row has alternating + buttons and thumbnails
        # Pattern: [+] [Page] [+] [Page] [+] [Page] [+]
        cols = 11  # 5 pages + 6 insert buttons per row
        page_cols = [1, 3, 5, 7, 9]  # Columns for pages
        insert_cols = [0, 2, 4, 6, 8, 10]  # Columns for + buttons
        
        pages_per_row = 5
        current_row = 0
        
        # Add initial "+" button at position 0
        self.create_insert_button(current_row, 0, 0)
        
        for i, page_data in enumerate(self.pages):
            # Calculate position
            page_in_row = i % pages_per_row
            current_row = i // pages_per_row
            col = page_cols[page_in_row]
            
            # Generate thumbnail
            thumbnail = self.generate_thumbnail(page_data)
            page_data['thumbnail'] = thumbnail
            
            # Create container frame for thumbnail
            container = tk.Frame(self.scrollable_frame, bg="white")
            container.grid(row=current_row, column=col, padx=2, pady=2, sticky="nsew")
            
            # Create frame for each thumbnail with border
            thumb_frame = tk.Frame(container, bg="white", relief=tk.SOLID, bd=2, 
                                  highlightthickness=0, cursor="hand2")
            thumb_frame.pack(fill=tk.BOTH, expand=True, padx=3, pady=3)
            
            # Store reference
            self.thumbnail_frames.append(thumb_frame)
            
            # Thumbnail label
            thumb_label = tk.Label(thumb_frame, image=thumbnail, bg="white", cursor="hand2")
            thumb_label.pack(padx=5, pady=5)
            
            # Page number badge
            page_num_label = tk.Label(thumb_frame, text=f"#{i+1}", bg="#2d8cf0", 
                                     fg="white", font=("Segoe UI", 9, "bold"), 
                                     padx=5, pady=2)
            page_num_label.pack()
            
            # Page info label
            if page_data['type'] == 'pdf':
                info_text = f"{os.path.basename(page_data['pdf_path'])}\nPage {page_data['page_num'] + 1}"
            else:
                info_text = f"{os.path.basename(page_data['image_path'])}"
            
            info_label = tk.Label(thumb_frame, text=info_text, bg="white", 
                                 font=("Segoe UI", 7), wraplength=140, cursor="hand2")
            info_label.pack(pady=(0, 5))
            
            # Bind drag and drop events to all widgets in the frame
            self.bind_drag_events(thumb_frame, i)
            self.bind_drag_events(thumb_label, i)
            self.bind_drag_events(info_label, i)
            self.bind_drag_events(page_num_label, i)
            
            # Add "+" button after this page
            insert_col = insert_cols[page_in_row + 1]
            self.create_insert_button(current_row, insert_col, i + 1)
        
        # Configure grid weights for responsive layout
        for col in range(cols):
            if col in page_cols:
                self.scrollable_frame.columnconfigure(col, weight=3, minsize=170)
            else:
                self.scrollable_frame.columnconfigure(col, weight=1, minsize=50)
    
    def create_insert_button(self, row, col, insert_position):
        """Create a '+' button for inserting pages"""
        # Create a small frame for the insert button
        insert_frame = tk.Frame(self.scrollable_frame, bg="white")
        insert_frame.grid(row=row, column=col, sticky="nsew", padx=2, pady=2)
        
        # Create the '+' button with better styling
        plus_btn = tk.Label(insert_frame, text="+", font=("Segoe UI", 20, "bold"),
                           bg="#e8e8e8", fg="#2d8cf0", cursor="hand2",
                           relief=tk.RAISED, bd=2, width=2, height=8)
        plus_btn.pack(expand=True, fill=tk.BOTH, padx=5, pady=5)
        
        # Bind click event
        plus_btn.bind("<Button-1>", lambda e: self.insert_at_position(insert_position))
        
        # Hover effects
        plus_btn.bind("<Enter>", lambda e: plus_btn.config(bg="#2d8cf0", fg="white", relief=tk.RAISED, bd=3))
        plus_btn.bind("<Leave>", lambda e: plus_btn.config(bg="#e8e8e8", fg="#2d8cf0", relief=tk.RAISED, bd=2))
    
    def bind_drag_events(self, widget, page_index):
        """Bind drag and drop events to a widget"""
        widget.bind("<Button-1>", lambda e: self.on_click(e, page_index))
        widget.bind("<B1-Motion>", lambda e: self.on_drag(e, page_index))
        widget.bind("<ButtonRelease-1>", lambda e: self.on_drop(e, page_index))
    
    def on_click(self, event, page_index):
        """Handle mouse click on thumbnail"""
        self.selected_page = page_index
        self.drag_data["x"] = event.x_root
        self.drag_data["y"] = event.y_root
        self.drag_data["item"] = page_index
        
        # Highlight selected thumbnail
        self.highlight_selected(page_index)
    
    def on_drag(self, event, page_index):
        """Handle dragging motion"""
        if self.drag_data["item"] == page_index:
            # Calculate drag distance
            dx = event.x_root - self.drag_data["x"]
            dy = event.y_root - self.drag_data["y"]
            
            # Check if we've moved enough to consider it a drag
            if abs(dx) > 10 or abs(dy) > 10:
                # Find which thumbnail we're hovering over
                widget_under_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
                
                # Visual feedback - highlight potential drop target
                for i, frame in enumerate(self.thumbnail_frames):
                    if widget_under_mouse in [frame] + list(frame.winfo_children()):
                        if i != page_index:
                            frame.config(bg="#ffeb3b", relief=tk.SOLID, bd=3)
                    else:
                        if i == self.selected_page:
                            frame.config(bg="#2d8cf0", relief=tk.SOLID, bd=3)
                        else:
                            frame.config(bg="white", relief=tk.SOLID, bd=2)
    
    def on_drop(self, event, page_index):
        """Handle drop event"""
        if self.drag_data["item"] is not None:
            # Find which thumbnail we dropped on
            widget_under_mouse = event.widget.winfo_containing(event.x_root, event.y_root)
            
            drop_index = None
            for i, frame in enumerate(self.thumbnail_frames):
                if widget_under_mouse in [frame] + list(frame.winfo_children()):
                    drop_index = i
                    break
            
            # Reorder pages if we found a valid drop target
            if drop_index is not None and drop_index != self.drag_data["item"]:
                dragged_page = self.pages.pop(self.drag_data["item"])
                self.pages.insert(drop_index, dragged_page)
                self.selected_page = drop_index
                
                # Refresh display
                self.refresh_thumbnails()
            else:
                # Just refresh highlighting
                self.highlight_selected(self.selected_page)
        
        self.drag_data["item"] = None
    
    def highlight_selected(self, page_index):
        """Highlight the selected thumbnail"""
        for i, frame in enumerate(self.thumbnail_frames):
            if i == page_index:
                frame.config(bg="#2d8cf0", relief=tk.SOLID, bd=3)
                # Update all child widgets background
                for child in frame.winfo_children():
                    if isinstance(child, tk.Label) and child.cget("text") != f"#{i+1}":
                        child.config(bg="#e3f2fd")
            else:
                frame.config(bg="white", relief=tk.SOLID, bd=2)
                for child in frame.winfo_children():
                    if isinstance(child, tk.Label) and child.cget("text") != f"#{i+1}":
                        child.config(bg="white")
    
    def insert_at_position(self, position):
        """Show dialog to insert image or PDF at specific position"""
        dialog = tk.Toplevel(self.window)
        dialog.title("Insert Content")
        dialog.geometry("300x150")
        dialog.transient(self.window)
        dialog.grab_set()
        
        tk.Label(dialog, text=f"Insert at position {position}:", 
                font=("Segoe UI", 11, "bold")).pack(pady=20)
        
        button_frame = tk.Frame(dialog)
        button_frame.pack(pady=10)
        
        def insert_image():
            dialog.destroy()
            self.add_image_at_position(position)
        
        def insert_pdf():
            dialog.destroy()
            self.add_pdf_at_position(position)
        
        ttk.Button(button_frame, text="Insert Image", command=insert_image).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Insert PDF", command=insert_pdf).pack(side=tk.LEFT, padx=10)
        ttk.Button(button_frame, text="Cancel", command=dialog.destroy).pack(side=tk.LEFT, padx=10)
    
    def add_image(self):
        """Add an image file as a new page"""
        position = self.ask_insert_position()
        if position is not None:
            self.add_image_at_position(position)
    
    def add_image_at_position(self, position):
        """Add an image file at a specific position"""
        file_path = filedialog.askopenfilename(
            title="Select Image",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.gif *.bmp *.tiff")]
        )
        
        if file_path:
            page_data = {
                'image_path': file_path,
                'type': 'image',
                'thumbnail': None
            }
            self.pages.insert(position, page_data)
            self.selected_page = position
            self.refresh_thumbnails()
    
    def add_pdf(self):
        """Add pages from another PDF file"""
        position = self.ask_insert_position()
        if position is not None:
            self.add_pdf_at_position(position)
    
    def add_pdf_at_position(self, position):
        """Add pages from another PDF file at a specific position"""
        file_path = filedialog.askopenfilename(
            title="Select PDF",
            filetypes=[("PDF files", "*.pdf")]
        )
        
        if file_path:
            try:
                doc = fitz.open(file_path)
                # Add all pages from the PDF
                for page_num in range(len(doc)):
                    page_data = {
                        'pdf_path': file_path,
                        'page_num': page_num,
                        'type': 'pdf',
                        'doc': doc,
                        'thumbnail': None
                    }
                    self.pages.insert(position + page_num, page_data)
                self.selected_page = position
                self.refresh_thumbnails()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")
    
    def ask_insert_position(self):
        """Ask user where to insert new content"""
        if not self.pages:
            return 0
        
        # Simple dialog for position
        position_dialog = tk.Toplevel(self.window)
        position_dialog.title("Insert Position")
        position_dialog.geometry("300x150")
        position_dialog.transient(self.window)
        position_dialog.grab_set()
        
        result = {"position": None}
        
        tk.Label(position_dialog, text="Insert at position:").pack(pady=10)
        
        position_var = tk.IntVar(value=len(self.pages))
        position_spinbox = tk.Spinbox(position_dialog, from_=0, to=len(self.pages), 
                                     textvariable=position_var, width=10)
        position_spinbox.pack(pady=5)
        
        button_frame = tk.Frame(position_dialog)
        button_frame.pack(pady=10)
        
        def ok_clicked():
            result["position"] = position_var.get()
            position_dialog.destroy()
        
        def cancel_clicked():
            position_dialog.destroy()
        
        ttk.Button(button_frame, text="OK", command=ok_clicked).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Cancel", command=cancel_clicked).pack(side=tk.LEFT, padx=5)
        
        position_dialog.wait_window()
        return result["position"]
    
    def delete_selected(self):
        """Delete the selected page"""
        if self.selected_page is not None and 0 <= self.selected_page < len(self.pages):
            if messagebox.askyesno("Confirm Delete", "Delete selected page?"):
                self.pages.pop(self.selected_page)
                self.selected_page = None
                self.refresh_thumbnails()
        else:
            messagebox.showwarning("Warning", "Please select a page to delete.")
    
    def save_pdf(self):
        """Save the edited PDF"""
        if not self.pages:
            messagebox.showwarning("Warning", "No pages to save.")
            return
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Save Edited PDF As"
        )
        
        if output_path:
            try:
                # First, close all PyMuPDF documents to release file locks
                opened_docs = {}
                for page_data in self.pages:
                    if page_data['type'] == 'pdf' and 'doc' in page_data:
                        pdf_path = page_data['pdf_path']
                        if pdf_path not in opened_docs:
                            opened_docs[pdf_path] = page_data['doc']
                
                # Close all documents
                for doc in opened_docs.values():
                    try:
                        doc.close()
                    except:
                        pass
                
                # Now create new PDF using PyPDF2
                writer = PdfWriter()
                temp_files = []
                
                for page_data in self.pages:
                    if page_data['type'] == 'pdf':
                        # Open PDF with PyPDF2 (files are now closed from PyMuPDF)
                        reader = PdfReader(page_data['pdf_path'])
                        writer.add_page(reader.pages[page_data['page_num']])
                    elif page_data['type'] == 'image':
                        # Convert image to PDF page and add
                        img = Image.open(page_data['image_path'])
                        
                        # Create temporary PDF from image
                        temp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
                        temp_pdf_path = temp_pdf.name
                        temp_pdf.close()
                        temp_files.append(temp_pdf_path)
                        
                        # Save image as PDF
                        img_rgb = img.convert('RGB')
                        img_rgb.save(temp_pdf_path, 'PDF')
                        
                        # Read the temporary PDF and add its page
                        temp_reader = PdfReader(temp_pdf_path)
                        writer.add_page(temp_reader.pages[0])
                
                # Write the final PDF
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                # Clean up temporary files
                for temp_file in temp_files:
                    try:
                        os.unlink(temp_file)
                    except:
                        pass
                
                # Reopen PyMuPDF documents for continued editing (if user wants to stay)
                for pdf_path, doc in opened_docs.items():
                    try:
                        new_doc = fitz.open(pdf_path)
                        # Update all page_data references to use the new document
                        for page_data in self.pages:
                            if page_data['type'] == 'pdf' and page_data['pdf_path'] == pdf_path:
                                page_data['doc'] = new_doc
                    except:
                        pass
                
                messagebox.showinfo("Success", f"PDF saved successfully!\nSaved to: {output_path}")
                self.close_editor()
                
            except Exception as e:
                # Clean up temp files on error
                for temp_file in temp_files:
                    try:
                        os.unlink(temp_file)
                    except:
                        pass
                
                # Try to reopen documents even on error
                for pdf_path in opened_docs.keys():
                    try:
                        new_doc = fitz.open(pdf_path)
                        for page_data in self.pages:
                            if page_data['type'] == 'pdf' and page_data['pdf_path'] == pdf_path:
                                page_data['doc'] = new_doc
                    except:
                        pass
                
                messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")
    
    def close_editor(self):
        """Close the editor window"""
        # Close any open PDF documents
        for page_data in self.pages:
            if page_data['type'] == 'pdf' and 'doc' in page_data:
                try:
                    page_data['doc'].close()
                except:
                    pass
        
        self.window.destroy()


def main():
    root = tk.Tk()
    app = PDFToolApp(root)
    # Minimize the console window after creating the Tkinter window
    app.minimize_console()
    root.mainloop()


if __name__ == "__main__":
    main()