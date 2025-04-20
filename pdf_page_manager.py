"""
PDF Page Manager
Copyright (C) 2025 Eng. Paulo J. Figueira & Aloia Aerospace, Inc
All rights reserved.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import fitz  # PyMuPDF
import os
from PIL import Image, ImageTk

class PDFManager:  # Level 0
    def __init__(self, root):  # Level 1
        self.root = root
        self.setup_window()
        self.pdf_path = None
        self.doc = None
        self.current_page = 0
        self.total_pages = 0
        self.rotations = {}
        self.pages_to_delete = set()
        self.create_widgets()

    def setup_window(self):  # Level 1
        self.root.title("PDF Manager v1.0 - by Paulo Victor")
        window_width = 800
        window_height = 600
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        center_x = int(screen_width/2 - window_width/2)
        center_y = int(screen_height/2 - window_height/2)
        self.root.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.root.resizable(True, True)

    def create_widgets(self):  # Level 1
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Left panel for PDF preview
        left_panel = ttk.Frame(main_frame)
        left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Preview canvas
        self.canvas = tk.Canvas(left_panel, bg='gray90')
        self.canvas.pack(fill=tk.BOTH, expand=True)

        # Right panel for controls
        right_panel = ttk.Frame(main_container, padding="10", width=200)
        right_panel.pack(side=tk.RIGHT, fill=tk.Y)
        right_panel.pack_propagate(False)

        # File operations
        file_frame = ttk.LabelFrame(right_panel, text="File Operations", padding="10")
        file_frame.pack(fill=tk.X, pady=5)

        ttk.Button(file_frame, text="Open PDF", command=self.open_pdf).pack(fill=tk.X, pady=2)
        ttk.Button(file_frame, text="Save Changes", command=self.save_changes).pack(fill=tk.X, pady=2)
        ttk.Separator(file_frame, orient='horizontal').pack(fill=tk.X, pady=5)
        ttk.Button(file_frame, text="Quit", command=self.quit_program, style='Quit.TButton').pack(fill=tk.X, pady=2)

        # Page navigation
        nav_frame = ttk.LabelFrame(right_panel, text="Page Navigation", padding="10")
        nav_frame.pack(fill=tk.X, pady=5)

        nav_buttons = ttk.Frame(nav_frame)
        nav_buttons.pack(fill=tk.X)
        ttk.Button(nav_buttons, text="Previous", command=self.prev_page).pack(side=tk.LEFT, padx=2)
        ttk.Button(nav_buttons, text="Next", command=self.next_page).pack(side=tk.LEFT, padx=2)

        self.page_label = ttk.Label(nav_frame, text="Page: 0/0")
        self.page_label.pack(pady=5)

        # Page operations
        ops_frame = ttk.LabelFrame(right_panel, text="Page Operations", padding="10")
        ops_frame.pack(fill=tk.X, pady=5)

        ttk.Button(ops_frame, text="Rotate Left (-90°)", command=lambda: self.rotate_page(-90)).pack(fill=tk.X, pady=2)
        ttk.Button(ops_frame, text="Rotate Right (+90°)", command=lambda: self.rotate_page(90)).pack(fill=tk.X, pady=2)
        ttk.Button(ops_frame, text="Mark for Deletion", command=self.toggle_delete_page).pack(fill=tk.X, pady=2)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_label = ttk.Label(right_panel, textvariable=self.status_var)
        status_label.pack(side=tk.BOTTOM, pady=5)

        # Style for quit button
        style = ttk.Style()
        style.configure('Quit.TButton', foreground='red')
        def quit_program(self):  # Level 1
        if hasattr(self, 'doc') and self.doc:
            self.doc.close()
        self.root.quit()
        self.root.destroy()

    def open_pdf(self):  # Level 1
        try:
            file_path = filedialog.askopenfilename(
                title="Select PDF File",
                filetypes=[("PDF files", "*.pdf")]
            )
            if file_path:
                if hasattr(self, 'doc') and self.doc:
                    self.doc.close()
                self.pdf_path = file_path
                self.doc = fitz.open(self.pdf_path)
                self.total_pages = len(self.doc)
                self.current_page = 0
                self.rotations = {}
                self.pages_to_delete = set()
                self.update_page_display()
                self.status_var.set(f"Loaded PDF with {self.total_pages} pages")
        except Exception as e:
            self.status_var.set("Error loading PDF")
            messagebox.showerror("Error", f"Failed to load PDF: {str(e)}")

    def update_page_display(self):  # Level 1
        if not self.doc:
            return

        self.page_label.config(text=f"Page: {self.current_page + 1}/{self.total_pages}")
        
        # Get the page
        page = self.doc[self.current_page]
        
        # Get the rotation angle for this page
        total_rotation = self.rotations.get(self.current_page, 0)
        
        # Create matrix for rotation
        matrix = fitz.Matrix(1, 1).prerotate(total_rotation)
        
        # Get the pixmap
        pix = page.get_pixmap(matrix=matrix)
        
        # Convert to PhotoImage
        img = tk.PhotoImage(data=pix.tobytes("ppm"))
        
        # Update canvas
        self.canvas.delete("all")
        # Store image reference to prevent garbage collection
        self.current_image = img
        
        # Calculate scaling to fit the canvas
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        scale_x = canvas_width / pix.width
        scale_y = canvas_height / pix.height
        scale = min(scale_x, scale_y)
        
        # Calculate centered position
        x = (canvas_width - pix.width * scale) / 2
        y = (canvas_height - pix.height * scale) / 2
        
        # Create image on canvas
        self.canvas.create_image(x, y, anchor=tk.NW, image=img)
        
        # If page is marked for deletion, overlay a red transparent rectangle
        if self.current_page in self.pages_to_delete:
            self.canvas.create_rectangle(0, 0, canvas_width, canvas_height,
                                      fill='red', stipple='gray50')

    def prev_page(self):  # Level 1
        if self.doc and self.current_page > 0:
            self.current_page -= 1
            self.update_page_display()

    def next_page(self):  # Level 1
        if self.doc and self.current_page < self.total_pages - 1:
            self.current_page += 1
            self.update_page_display()

    def rotate_page(self, angle):  # Level 1
        if not self.doc:
            return
            
        current_rotation = self.rotations.get(self.current_page, 0)
        new_rotation = (current_rotation + angle) % 360
        self.rotations[self.current_page] = new_rotation
        self.update_page_display()
        self.status_var.set(f"Page {self.current_page + 1} rotated by {angle}°")

    def toggle_delete_page(self):  # Level 1
        if not self.doc:
            return
            
        if self.current_page in self.pages_to_delete:
            self.pages_to_delete.remove(self.current_page)
            self.status_var.set(f"Page {self.current_page + 1} unmarked for deletion")
        else:
            self.pages_to_delete.add(self.current_page)
            self.status_var.set(f"Page {self.current_page + 1} marked for deletion")
        self.update_page_display()

    def save_changes(self):  # Level 1
        if not self.doc:
            return
            
        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")]
            )
            if not output_path:
                return

            # Check if file exists and try to remove it
            if os.path.exists(output_path):
                try:
                    # Close the existing document if it's the same one we have open
                    if hasattr(self, 'doc'):
                        self.doc.close()
                    os.remove(output_path)
                except PermissionError:
                    # If removal fails, try saving to a temporary file
                    temp_path = output_path + ".temp"
                    self.save_to_path(temp_path)
                    try:
                        os.remove(output_path)
                        os.rename(temp_path, output_path)
                    except Exception as e:
                        messagebox.showerror("Error", 
                            "Cannot overwrite existing file. Try saving with a different name.")
                        return
            else:
                # Save directly if file doesn't exist
                self.save_to_path(output_path)

            # Reload the saved PDF
            self.doc = fitz.open(output_path)
            self.total_pages = len(self.doc)
            self.current_page = min(self.current_page, self.total_pages - 1)
            self.rotations = {}
            self.pages_to_delete = set()
            self.update_page_display()
            
            self.status_var.set("Changes saved successfully")
            messagebox.showinfo("Success", "PDF saved successfully!")
            
        except Exception as e:
            self.status_var.set("Error saving PDF")
            messagebox.showerror("Error", f"Failed to save PDF: {str(e)}")

    def save_to_path(self, output_path):  # Level 1
        output_doc = fitz.open()
        
        # Copy pages to new document
        for page_num in range(self.total_pages):
            if page_num not in self.pages_to_delete:
                page = self.doc[page_num]
                if page_num in self.rotations:
                    rotation = self.rotations[page_num]
                    # Create new page
                    new_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
                    # Apply rotation
                    new_page.set_rotation(rotation)
                    new_page.show_pdf_page(new_page.rect, self.doc, page_num)
                else:
                    output_doc.insert_pdf(self.doc, from_page=page_num, to_page=page_num)

        # Save the modified PDF
        output_doc.save(output_path)
        output_doc.close()

def main():  # Level 0
    try:
        root = tk.Tk()
        app = PDFManager(root)
        root.mainloop()
    except Exception as e:
        print(f"Error in main: {str(e)}")
        messagebox.showerror("Error", f"Application error: {str(e)}")

if __name__ == "__main__":
    main()
    app = PDFManager(root)
    root.mainloop()


if __name__ == "__main__":
    main()
