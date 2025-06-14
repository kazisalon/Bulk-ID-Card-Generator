import os
import re
import logging
from PIL import Image, ImageDraw, ImageFont, ImageTk
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

class IDCardGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("ID Card Generator - Dark Mode")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 800)
        
        # Configure dark theme
        self.setup_dark_theme()
        
        # Initialize paths
        self.template_path = tk.StringVar()
        self.photos_folder = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.font_path = tk.StringVar()
        self.font_size = tk.StringVar(value="20")
        
        # Template customization options
        self.photo_frame_style = tk.StringVar(value="circle")
        self.font_color = tk.StringVar(value="black")
        self.border_size = tk.StringVar(value="2")  # Default border size of 2 pixels
        
        # Initialize preview variables
        self.preview_image = None
        self.current_coordinate = None
        self.coordinate_labels = ["Photo", "QR Code", "Name", "Class", "Contact", "Address", "Guardian", "Validity", "Roll No.", "RegNo"]
        self.current_label_index = 0
        self.coordinates = {}
        self.scale_factor = 1.0
        
        # Initialize paths for QR codes
        self.qr_folder = tk.StringVar()
        
        # PDF Export Option
        self.export_as_pdf = tk.BooleanVar(value=False)
        
        self.create_widgets()
        
    def setup_dark_theme(self):
        """Configure dark theme colors and styles"""
        # Configure ttk styles for dark theme
        style = ttk.Style()
        
        # Configure dark theme colors
        self.colors = {
            'bg_primary': '#1e1e1e',       # Main background
            'bg_secondary': '#2d2d2d',     # Secondary background
            'bg_tertiary': '#3d3d3d',      # Tertiary background
            'fg_primary': '#ffffff',        # Primary text
            'fg_secondary': '#cccccc',      # Secondary text
            'accent': '#0078d4',           # Accent color
            'accent_hover': '#106ebe',     # Accent hover
            'success': '#107c10',          # Success color
            'warning': '#ff8c00',          # Warning color
            'error': '#d13438',            # Error color
            'border': '#484848'            # Border color
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['bg_primary'])
        
        # Configure ttk styles
        style.theme_use('clam')
        
        # Configure Frame styles
        style.configure('Dark.TFrame', 
                       background=self.colors['bg_primary'],
                       borderwidth=0)
        
        style.configure('Card.TFrame',
                       background=self.colors['bg_secondary'],
                       relief='solid',
                       borderwidth=1)
        
        # Configure Label styles
        style.configure('Dark.TLabel',
                       background=self.colors['bg_primary'],
                       foreground=self.colors['fg_primary'],
                       font=('Segoe UI', 10))
        
        style.configure('Heading.TLabel',
                       background=self.colors['bg_primary'],
                       foreground=self.colors['fg_primary'],
                       font=('Segoe UI', 12, 'bold'))
        
        # Configure Entry styles
        style.configure('Dark.TEntry',
                       fieldbackground=self.colors['bg_tertiary'],
                       borderwidth=1,
                       insertcolor=self.colors['fg_primary'],
                       foreground=self.colors['fg_primary'])
        
        # Configure Button styles
        style.configure('Dark.TButton',
                       background=self.colors['accent'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 9))
        
        style.map('Dark.TButton',
                 background=[('active', self.colors['accent_hover']),
                           ('pressed', self.colors['accent_hover'])])
        
        style.configure('Success.TButton',
                       background=self.colors['success'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       font=('Segoe UI', 10, 'bold'))
        
        style.map('Success.TButton',
                 background=[('active', '#0e6e0e'),
                           ('pressed', '#0e6e0e')])
        
        # Configure LabelFrame styles
        style.configure('Dark.TLabelframe',
                       background=self.colors['bg_secondary'],
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Dark.TLabelframe.Label',
                       background=self.colors['bg_secondary'],
                       foreground=self.colors['fg_primary'],
                       font=('Segoe UI', 10, 'bold'))
        
    def create_widgets(self):
        # Create outer frame to hold canvas and scrollbar
        outer_frame = ttk.Frame(self.root, style='Dark.TFrame')
        outer_frame.pack(fill='both', expand=True)
        
        # Create canvas and scrollbar for vertical scrolling
        canvas = tk.Canvas(outer_frame, bg=self.colors['bg_primary'], highlightthickness=0)
        scrollbar = ttk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        
        # Create main container with padding
        main_container = ttk.Frame(canvas, style='Dark.TFrame', padding="20")
        
        # Configure scrolling
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Pack scrollbar and canvas
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        
        # Bind mouse wheel to parent frame for consistent scrolling
        outer_frame.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        
        # Create window in canvas
        canvas_window = canvas.create_window((0, 0), window=main_container, anchor="nw", tags="main_container")
        
        # Configure main window grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        main_container.rowconfigure(0, weight=1)
        
        # Configure canvas scrolling
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas.itemconfig(canvas_window, width=canvas.winfo_width())
            # Ensure minimum height for scrolling
            content_height = main_container.winfo_reqheight()
            canvas_height = canvas.winfo_height()
            canvas.configure(scrollregion=(0, 0, canvas.winfo_width(), max(content_height, canvas_height)))
            
        main_container.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", lambda e: configure_scroll_region())
        
        # Initial configuration
        self.root.update_idletasks()
        configure_scroll_region()
        
        # Create left panel (controls)
        left_panel = self.create_left_panel(main_container)
        left_panel.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 20))
        
        # Create right panel (preview and coordinates)
        right_panel = self.create_right_panel(main_container)
        right_panel.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
    def create_left_panel(self, parent):
        """Create the left control panel"""
        left_frame = ttk.Frame(parent, style='Dark.TFrame')
        left_frame.columnconfigure(1, weight=1)
        
        # File Selection Section
        file_section = ttk.LabelFrame(left_frame, text="File Selection", style='Dark.TLabelframe', padding="15")
        file_section.grid(row=0, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        file_section.columnconfigure(1, weight=1)
        
        # Template selection
        ttk.Label(file_section, text="ID Card Template:", style='Dark.TLabel').grid(row=0, column=0, sticky=tk.W, pady=8)
        ttk.Entry(file_section, textvariable=self.template_path, style='Dark.TEntry').grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(file_section, text="Browse", command=self.browse_template, style='Dark.TButton').grid(row=0, column=2)
        
        # Photos folder selection
        ttk.Label(file_section, text="Photos Folder:", style='Dark.TLabel').grid(row=1, column=0, sticky=tk.W, pady=8)
        ttk.Entry(file_section, textvariable=self.photos_folder, style='Dark.TEntry').grid(row=1, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(file_section, text="Browse", command=lambda: self.browse_folder(self.photos_folder), style='Dark.TButton').grid(row=1, column=2)
        
        # QR codes folder selection
        ttk.Label(file_section, text="QR Codes Folder:", style='Dark.TLabel').grid(row=2, column=0, sticky=tk.W, pady=8)
        ttk.Entry(file_section, textvariable=self.qr_folder, style='Dark.TEntry').grid(row=2, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(file_section, text="Browse", command=lambda: self.browse_folder(self.qr_folder), style='Dark.TButton').grid(row=2, column=2)
        
        # Excel file selection
        ttk.Label(file_section, text="Excel File:", style='Dark.TLabel').grid(row=3, column=0, sticky=tk.W, pady=8)
        ttk.Entry(file_section, textvariable=self.excel_path, style='Dark.TEntry').grid(row=3, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(file_section, text="Browse", command=lambda: self.browse_file(self.excel_path, [("Excel files", "*.xlsx")]), style='Dark.TButton').grid(row=3, column=2)
        
        # Output folder selection
        ttk.Label(file_section, text="Output Folder:", style='Dark.TLabel').grid(row=4, column=0, sticky=tk.W, pady=8)
        ttk.Entry(file_section, textvariable=self.output_folder, style='Dark.TEntry').grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(file_section, text="Browse", command=lambda: self.browse_folder(self.output_folder), style='Dark.TButton').grid(row=4, column=2)
        
        # PDF Export Option
        pdf_export_checkbox = ttk.Checkbutton(file_section, text="Export as single PDF (A4 Landscape)", variable=self.export_as_pdf, style='Dark.TLabel') # Using Dark.TLabel style for text color
        pdf_export_checkbox.grid(row=5, column=0, columnspan=3, sticky=tk.W, pady=(10, 0))
        
        # Font Configuration Section
        font_section = ttk.LabelFrame(left_frame, text="Font Configuration", style='Dark.TLabelframe', padding="15")
        font_section.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        font_section.columnconfigure(1, weight=1)
        
        # Font selection
        ttk.Label(font_section, text="Font File (optional):", style='Dark.TLabel').grid(row=0, column=0, sticky=tk.W, pady=8)
        ttk.Entry(font_section, textvariable=self.font_path, style='Dark.TEntry').grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 10))
        ttk.Button(font_section, text="Browse", command=lambda: self.browse_file(self.font_path, [("Font files", "*.ttf;*.otf")]), style='Dark.TButton').grid(row=0, column=2)
        
        # Font sizes for different fields
        self.font_sizes = {}
        for i, label in enumerate(self.coordinate_labels):
            if label not in ['Photo', 'QR Code']:  # Skip non-text elements
                ttk.Label(font_section, text=f"{label} Font Size:", style='Dark.TLabel').grid(row=i+1, column=0, sticky=tk.W, pady=5)
                size_var = tk.StringVar(value="20")  # Default size
                self.font_sizes[label] = size_var
                size_entry = ttk.Entry(font_section, textvariable=size_var, width=8, style='Dark.TEntry')
                size_entry.grid(row=i+1, column=1, sticky=tk.W, padx=(10, 0))
        
        # Template Customization Section
        template_section = ttk.LabelFrame(left_frame, text="Template Customization", style='Dark.TLabelframe', padding="15")
        template_section.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        template_section.columnconfigure(1, weight=1)
        
        # Photo Frame Style
        ttk.Label(template_section, text="Photo Frame Style:", style='Dark.TLabel').grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Radiobutton(template_section, text="Circle", variable=self.photo_frame_style, value="circle",
                       style='Dark.TRadiobutton').grid(row=0, column=1, sticky=tk.W, padx=5)
        ttk.Radiobutton(template_section, text="Square", variable=self.photo_frame_style, value="square",
                       style='Dark.TRadiobutton').grid(row=0, column=2, sticky=tk.W, padx=5)
        
        # Border Size
        ttk.Label(template_section, text="Border Size (px):", style='Dark.TLabel').grid(row=1, column=0, sticky=tk.W, pady=5)
        border_size_entry = ttk.Entry(template_section, textvariable=self.border_size, width=8, style='Dark.TEntry')
        border_size_entry.grid(row=1, column=1, sticky=tk.W, padx=5)
        
        # Font Color Selection
        ttk.Label(template_section, text="Font Color:", style='Dark.TLabel').grid(row=2, column=0, sticky=tk.W, pady=5)
        color_options = ["black", "blue", "red", "green", "#235cca"]  # Added #0080c2
        color_dropdown = ttk.Combobox(template_section, textvariable=self.font_color, values=color_options, state="readonly")
        color_dropdown.grid(row=2, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5)
        
        # Border Color Selection
        self.border_color = tk.StringVar(value="blue") # Default to blue border
        ttk.Label(template_section, text="Border Color:", style='Dark.TLabel').grid(row=3, column=0, sticky=tk.W, pady=5)
        border_color_options = ["black", "blue", "red", "green", "white", "#235cca"]
        border_color_dropdown = ttk.Combobox(template_section, textvariable=self.border_color, values=border_color_options, state="readonly")
        border_color_dropdown.grid(row=3, column=1, columnspan=2, sticky=(tk.W, tk.E), padx=5)
        
        # Configure style for combobox
        style = ttk.Style()
        style.configure('TCombobox',
                      fieldbackground=self.colors['bg_tertiary'],
                      background=self.colors['bg_secondary'],
                      foreground=self.colors['fg_primary'])
        color_dropdown.configure(style='TCombobox')
        
        # Generate Button
        generate_frame = ttk.Frame(left_frame, style='Dark.TFrame')
        generate_frame.grid(row=4, column=0, columnspan=3, pady=20)
        ttk.Button(generate_frame, text="🚀 Generate ID Cards", command=self.generate_cards, 
                  style='Success.TButton', width=25).pack()
        
        # Log Section
        log_section = ttk.LabelFrame(left_frame, text="Generation Log", style='Dark.TLabelframe', padding="15")
        log_section.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        log_section.columnconfigure(0, weight=1)
        log_section.rowconfigure(0, weight=1)
        
        # Log display with dark theme
        self.log_text = ScrolledText(log_section, width=70, height=12, 
                                   bg=self.colors['bg_tertiary'], 
                                   fg=self.colors['fg_primary'],
                                   insertbackground=self.colors['fg_primary'],
                                   selectbackground=self.colors['accent'],
                                   font=('Consolas', 9))
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure left frame row weights
        left_frame.rowconfigure(6, weight=1)
        
        return left_frame
        
    def create_right_panel(self, parent):
        """Create the right preview panel"""
        right_frame = ttk.Frame(parent, style='Dark.TFrame')
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(1, weight=1)
        
        # Preview Section
        preview_section = ttk.LabelFrame(right_frame, text="Template Preview", style='Dark.TLabelframe', padding="15")
        preview_section.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        preview_section.columnconfigure(0, weight=1)
        
        # Preview canvas with dark theme
        canvas_frame = ttk.Frame(preview_section, style='Dark.TFrame')
        canvas_frame.grid(row=0, column=0)
        
        self.preview_canvas = tk.Canvas(canvas_frame, width=400, height=500, 
                                      bg=self.colors['bg_tertiary'], 
                                      highlightthickness=1,
                                      highlightbackground=self.colors['border'])
        self.preview_canvas.pack()
        
        # Initialize drag and drop variables
        self.drag_data = {"x": 0, "y": 0, "item": None, "label": None}
        self.draggable_items = {}
        
        # Bind drag and drop events
        self.preview_canvas.tag_bind("draggable", "<ButtonPress-1>", self.on_drag_start)
        self.preview_canvas.tag_bind("draggable", "<ButtonRelease-1>", self.on_drag_stop)
        self.preview_canvas.tag_bind("draggable", "<B1-Motion>", self.on_drag_motion)
        
        # Current coordinate info
        coord_info_frame = ttk.Frame(preview_section, style='Dark.TFrame')
        coord_info_frame.grid(row=1, column=0, pady=(15, 0))
        
        ttk.Label(coord_info_frame, text="Drag elements to set their positions:", style='Dark.TLabel').grid(row=0, column=0, padx=(0, 10))
        
        # Add a button to create draggable elements
        ttk.Button(coord_info_frame, text="Create Draggable Elements", 
                  command=self.create_draggable_elements,
                  style='Dark.TButton').grid(row=0, column=1)
        
        # Coordinates Section
        coord_section = ttk.LabelFrame(right_frame, text="Coordinate Management", style='Dark.TLabelframe', padding="15")
        coord_section.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        coord_section.columnconfigure(0, weight=1)
        coord_section.rowconfigure(0, weight=1)
        
        # Create notebook for coordinates
        coord_notebook = ttk.Notebook(coord_section)
        coord_notebook.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        
        # Manual input tab
        manual_frame = ttk.Frame(coord_notebook, style='Dark.TFrame', padding="10")
        coord_notebook.add(manual_frame, text="Manual Input")
        
        # Create scrollable frame for manual coordinates
        manual_canvas = tk.Canvas(manual_frame, bg=self.colors['bg_secondary'], highlightthickness=0)
        manual_scrollbar = ttk.Scrollbar(manual_frame, orient="vertical", command=manual_canvas.yview)
        manual_scrollable_frame = ttk.Frame(manual_canvas, style='Dark.TFrame')
        
        manual_scrollable_frame.bind(
            "<Configure>",
            lambda e: manual_canvas.configure(scrollregion=manual_canvas.bbox("all"))
        )
        
        manual_canvas.create_window((0, 0), window=manual_scrollable_frame, anchor="nw")
        manual_canvas.configure(yscrollcommand=manual_scrollbar.set)
        
        manual_canvas.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        manual_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        manual_frame.columnconfigure(0, weight=1)
        manual_frame.rowconfigure(0, weight=1)
        
        # Create entry fields for each coordinate
        self.coordinate_entries = {}
        for i, label in enumerate(self.coordinate_labels):
            # Create frame for each coordinate
            coord_item_frame = ttk.Frame(manual_scrollable_frame, style='Card.TFrame', padding="10")
            coord_item_frame.grid(row=i, column=0, sticky=(tk.W, tk.E), pady=5, padx=5)
            coord_item_frame.columnconfigure(1, weight=1)
            coord_item_frame.columnconfigure(3, weight=1)
            
            ttk.Label(coord_item_frame, text=f"{label}:", style='Dark.TLabel', width=12).grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            
            # X coordinate
            ttk.Label(coord_item_frame, text="X:", style='Dark.TLabel').grid(row=0, column=1, sticky=tk.E, padx=(0, 5))
            x_var = tk.StringVar()
            x_entry = ttk.Entry(coord_item_frame, textvariable=x_var, width=8, style='Dark.TEntry')
            x_entry.grid(row=0, column=2, padx=(0, 15))
            
            # Y coordinate
            ttk.Label(coord_item_frame, text="Y:", style='Dark.TLabel').grid(row=0, column=3, sticky=tk.E, padx=(0, 5))
            y_var = tk.StringVar()
            y_entry = ttk.Entry(coord_item_frame, textvariable=y_var, width=8, style='Dark.TEntry')
            y_entry.grid(row=0, column=4, padx=(0, 15))
            
            # Update button
            update_btn = ttk.Button(coord_item_frame, text="Set", 
                                  command=lambda l=label: self.update_coordinate_from_entry(l),
                                  style='Dark.TButton', width=8)
            update_btn.grid(row=0, column=5)
            
            # Store the variables
            self.coordinate_entries[label] = (x_var, y_var)
        
        manual_scrollable_frame.columnconfigure(0, weight=1)
        
        # Coordinates display tab
        display_frame = ttk.Frame(coord_notebook, style='Dark.TFrame', padding="10")
        coord_notebook.add(display_frame, text="Current Coordinates")
        
        self.coordinates_text = ScrolledText(display_frame, width=45, height=15,
                                           bg=self.colors['bg_tertiary'],
                                           fg=self.colors['fg_primary'],
                                           insertbackground=self.colors['fg_primary'],
                                           selectbackground=self.colors['accent'],
                                           font=('Consolas', 10))
        self.coordinates_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        display_frame.columnconfigure(0, weight=1)
        display_frame.rowconfigure(0, weight=1)
        
        return right_frame
        
    def browse_template(self):
        filename = filedialog.askopenfilename(
            title="Select ID Card Template",
            filetypes=[("PNG files", "*.png"), ("All image files", "*.png *.jpg *.jpeg")]
        )
        if filename:
            self.template_path.set(filename)
            self.load_preview_image()
    
    def browse_file(self, string_var, file_types):
        filename = filedialog.askopenfilename(filetypes=file_types)
        if filename:
            string_var.set(filename)
            
    def browse_folder(self, string_var):
        folder = filedialog.askdirectory()
        if folder:
            string_var.set(folder)
            
    def load_preview_image(self):
        if self.template_path.get():
            try:
                image = Image.open(self.template_path.get())
                # Get original size
                original_width, original_height = image.size
                
                # Calculate scale factor to fit in canvas
                canvas_width = 400
                canvas_height = 500
                width_scale = canvas_width / original_width
                height_scale = canvas_height / original_height
                self.scale_factor = min(width_scale, height_scale)
                
                # Resize image
                new_width = int(original_width * self.scale_factor)
                new_height = int(original_height * self.scale_factor)
                resized_image = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
                
                self.preview_image = ImageTk.PhotoImage(resized_image)
                self.preview_canvas.delete("all")
                self.preview_canvas.create_image(canvas_width//2, canvas_height//2, image=self.preview_image)
                
                # Reset coordinates
                self.coordinates = {}
                self.current_label_index = 0
                self.coordinates_text.delete(1.0, tk.END)
                
                self.log_message(f"✅ Template loaded: {os.path.basename(self.template_path.get())}")
            except Exception as e:
                messagebox.showerror("Error", f"Error loading template: {str(e)}")
                self.log_message(f"❌ Error loading template: {str(e)}")
    
    def create_draggable_elements(self):
        """Create draggable elements for each coordinate label."""
        if not self.preview_image:
            messagebox.showwarning("Warning", "Please load a template first!")
            return
            
        # Clear existing draggable items
        for item in self.draggable_items.values():
            self.preview_canvas.delete(item)
        self.draggable_items.clear()
        
        # Create draggable elements for each coordinate
        for label in self.coordinate_labels:
            # Create a draggable element (circle with label)
            x = 200  # Start in the middle of the canvas
            y = 250
            
            # Determine circle color based on label type
            fill_color = '#cccccc' # Default color for text labels
            if label == "Photo":
                fill_color = 'blue' # Color for Photo
            elif label == "QR Code":
                fill_color = self.colors['success'] # Color for QR Code

            # Create the circle
            circle = self.preview_canvas.create_oval(x-15, y-15, x+15, y+15,
                                                  fill=fill_color,
                                                  outline='white',
                                                  width=2,
                                                  tags=("draggable", f"element_{label}"))
            
            # Create the label
            text = self.preview_canvas.create_text(x, y,
                                                 text=label,
                                                 fill='red',
                                                 font=('Segoe UI', 9, 'bold'),
                                                 tags=("draggable", f"element_{label}"))
            
            # Store the items
            self.draggable_items[label] = (circle, text)
            
            # If we have existing coordinates, move the element there
            if label in self.coordinates:
                x, y = self.coordinates[label]
                scaled_x = int(x * self.scale_factor)
                scaled_y = int(y * self.scale_factor)
                self.move_draggable_element(label, scaled_x, scaled_y)
        
        self.log_message("✅ Created draggable elements. Drag them to set positions.")

    def on_drag_start(self, event):
        """Handle the start of a drag operation."""
        # Get the clicked item
        item = self.preview_canvas.find_closest(event.x, event.y)[0] # find_closest returns a tuple, get the first item
        if not item or "draggable" not in self.preview_canvas.gettags(item):
            # Clicked item is not a draggable element
            return

        # Find the label associated with the dragged item
        label = None
        for tag in self.preview_canvas.gettags(item):
            if tag.startswith("element_"):
                label = tag.replace("element_", "")
                break

        if not label or label not in self.draggable_items:
            # Could not find a valid label for the item
            return

        # Store the drag data: item, label, and its current canvas coordinates
        self.drag_data["item"] = item
        self.drag_data["label"] = label
        self.drag_data["start_x"] = event.x
        self.drag_data["start_y"] = event.y
        # Get the current position of the item (center of the circle/text)
        bbox = self.preview_canvas.bbox(item)
        if bbox:
            self.drag_data["item_start_x"] = (bbox[0] + bbox[2]) / 2
            self.drag_data["item_start_y"] = (bbox[1] + bbox[3]) / 2
        else:
             # Fallback if bbox is not available (shouldn't happen for valid items)
            self.drag_data["item_start_x"] = event.x
            self.drag_data["item_start_y"] = event.y

        # Change appearance to show it's being dragged
        # Find the circle element for the label
        circle_item = None
        for item_id in self.preview_canvas.find_withtag(f"element_{label}"):
             if 'oval' in self.preview_canvas.type(item_id):
                  circle_item = item_id
                  break

        if circle_item:
            self.preview_canvas.itemconfig(circle_item, fill=self.colors['accent_hover']) # Use accent hover color

    def on_drag_motion(self, event):
        """Handle the motion of a drag operation."""
        if not self.drag_data["item"]:
            return

        # Calculate the distance moved from the start of the drag
        dx = event.x - self.drag_data["start_x"]
        dy = event.y - self.drag_data["start_y"]

        # Calculate the new desired position based on the initial item position and total drag distance
        new_x = self.drag_data["item_start_x"] + dx
        new_y = self.drag_data["item_start_y"] + dy

        # Move the element (both circle and text) to the new position
        label = self.drag_data["label"]
        circle, text = self.draggable_items[label] # Assuming draggable_items stores both circle and text IDs

        # Move circle
        self.preview_canvas.coords(circle, new_x - 15, new_y - 15, new_x + 15, new_y + 15)
        # Move text
        self.preview_canvas.coords(text, new_x, new_y)


    def on_drag_stop(self, event):
        """Handle the end of a drag operation."""
        if not self.drag_data["item"]:
            return

        label = self.drag_data["label"]
        item = self.drag_data["item"]

        # Get the final canvas coordinates of the item
        # We need the center coordinate after the motion is complete
        bbox = self.preview_canvas.bbox(item)
        if not bbox:
             # Should not happen if item was valid
             self.log_message(f"❌ Error getting bbox for item {item} on drag stop.")
             self.drag_data = {"item": None, "label": None, "start_x": 0, "start_y": 0, "item_start_x": 0, "item_start_y": 0} # Reset drag data
             return

        # Calculate the center of the element on the canvas
        final_canvas_x = (bbox[0] + bbox[2]) / 2
        final_canvas_y = (bbox[1] + bbox[3]) / 2

        # Get the image position on the canvas. Assuming the image is centered,
        # its top-left corner is at (canvas_width//2 - image_width//2, canvas_height//2 - image_height//2).
        # However, since we scaled the image to fit, we need the *scaled* image dimensions.
        # The image is created with its center at (canvas_width//2, canvas_height//2).
        # So, the image origin on canvas is (canvas_width//2 - new_width//2, canvas_height//2 - new_height//2)
        # Where new_width = original_width * scale_factor, new_height = original_height * scale_factor

        # Let's calculate the image origin on the canvas
        canvas_width = self.preview_canvas.winfo_width() # Use current canvas size
        canvas_height = self.preview_canvas.winfo_height()

        if not self.preview_image:
             # Should not happen if template was loaded
             self.log_message("❌ Preview image not loaded on drag stop.")
             self.drag_data = {"item": None, "label": None, "start_x": 0, "start_y": 0, "item_start_x": 0, "item_start_y": 0} # Reset drag data
             return

        # Get the *displayed* image dimensions from the PhotoImage
        # Note: PhotoImage.width() and height() give the displayed size
        displayed_image_width = self.preview_image.width()
        displayed_image_height = self.preview_image.height()

        # Calculate the top-left corner of the centered image on the canvas
        image_canvas_origin_x = (canvas_width // 2) - (displayed_image_width // 2)
        image_canvas_origin_y = (canvas_height // 2) - (displayed_image_height // 2)

        # Calculate the coordinate relative to the image origin on the canvas
        relative_image_x = final_canvas_x - image_canvas_origin_x
        relative_image_y = final_canvas_y - image_canvas_origin_y

        # Scale the relative image coordinate back to the original image size
        if self.scale_factor > 0: # Avoid division by zero
             actual_x = int(relative_image_x / self.scale_factor)
             actual_y = int(relative_image_y / self.scale_factor)
        else:
             # Fallback if scale_factor is not valid
             actual_x = int(final_canvas_x) # Use canvas coordinate if scaling fails
             actual_y = int(final_canvas_y)
             self.log_message("⚠️ Invalid scale factor on drag stop. Using canvas coordinates.")


        # Ensure coordinates are non-negative
        actual_x = max(0, actual_x)
        actual_y = max(0, actual_y)

        # Store the coordinates
        self.coordinates[label] = (actual_x, actual_y)

        # Update manual entry fields
        if label in self.coordinate_entries:
             x_var, y_var = self.coordinate_entries[label]
             x_var.set(str(actual_x))
             y_var.set(str(actual_y))


        # Update the display
        self.update_coordinates_display()

        # Reset the appearance of the circle
        # Find the circle element for the label
        circle_item = None
        for item_id in self.preview_canvas.find_withtag(f"element_{label}"):
             if 'oval' in self.preview_canvas.type(item_id):
                  circle_item = item_id
                  break

        if circle_item:
            # Determine reset color based on label type
            reset_color = '#cccccc' # Default color for text labels
            if label == "Photo":
                 fill_color = 'blue' # Color for Photo
            elif label == "QR Code":
                 fill_color = self.colors['success'] # Use success color for QR

            if label != "Photo" and label != "QR Code":
                 # For text labels, use the secondary foreground color
                 reset_color = self.colors['fg_secondary']
            else:
                 # For Photo and QR Code, use their specific colors
                 reset_color = fill_color # Use the fill_color determined earlier

            self.preview_canvas.itemconfig(circle_item, fill=reset_color) # Reset to original color


        # Clear drag data
        self.drag_data = {"item": None, "label": None, "start_x": 0, "start_y": 0, "item_start_x": 0, "item_start_y": 0}

        self.log_message(f"📍 Set {label} coordinate: ({actual_x}, {actual_y})")

    def move_draggable_element(self, label, x, y):
        """Move a draggable element to the specified position."""
        circle, text = self.draggable_items[label]
        
        # Move the circle
        self.preview_canvas.coords(circle, x-15, y-15, x+15, y+15)
        
        # Move the text
        self.preview_canvas.coords(text, x, y)
    
    def update_coordinate_from_entry(self, label):
        """Update coordinate from manual entry fields."""
        try:
            x_var, y_var = self.coordinate_entries[label]
            x = int(x_var.get())
            y = int(y_var.get())
            
            # Store the coordinate
            self.coordinates[label] = (x, y)
            
            # Update the display
            self.update_coordinates_display()
            
            # Move the draggable element on the canvas to the scaled position
            scaled_x = int(x * self.scale_factor)
            scaled_y = int(y * self.scale_factor)
            
            # Calculate the position relative to the image origin on the canvas
            # Assuming the image is centered, find its top-left corner on canvas
            canvas_width = self.preview_canvas.winfo_width()
            canvas_height = self.preview_canvas.winfo_height()
            if not self.preview_image:
                self.log_message("❌ Preview image not loaded for manual coordinate update.")
                return

            displayed_image_width = self.preview_image.width()
            displayed_image_height = self.preview_image.height()

            image_canvas_origin_x = (canvas_width // 2) - (displayed_image_width // 2)
            image_canvas_origin_y = (canvas_height // 2) - (displayed_image_height // 2)

            # Calculate the target canvas coordinate based on the scaled coordinate relative to image origin
            target_canvas_x = image_canvas_origin_x + scaled_x
            target_canvas_y = image_canvas_origin_y + scaled_y


            # Find the draggable element for this label and move it
            if label in self.draggable_items:
                 circle, text = self.draggable_items[label]
                 self.move_draggable_element(label, target_canvas_x, target_canvas_y) # Pass target canvas coords


            # Clear previous points drawn by manual entry (if any)
            self.preview_canvas.delete(f"point_{label}")

            # Optional: Draw a temporary visual cue at the location set by manual input on the canvas
            # This might be confusing with the draggable elements now also moving, so let's skip for now.
            # If needed, we can add a different visual indicator.

            # Find the index of the current label and move to the next one
            try:
                current_index = self.coordinate_labels.index(label)
                self.current_label_index = (current_index + 1) % len(self.coordinate_labels)
            except ValueError:
                pass

            self.log_message(f"📍 Manually set {label} coordinate: ({x}, {y}). Updated preview element.")

        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for coordinates")
            self.log_message(f"❌ Invalid coordinate values for {label}")
    
    def update_coordinates_display(self):
        """Update the coordinates display text widget"""
        self.coordinates_text.delete(1.0, tk.END)
        if self.coordinates:
            self.coordinates_text.insert(tk.END, "Current Coordinates:\n" + "="*30 + "\n\n")
            for label, (x, y) in self.coordinates.items():
                self.coordinates_text.insert(tk.END, f"{label:12}: ({x:4}, {y:4})\n")
        else:
            self.coordinates_text.insert(tk.END, "No coordinates set yet.\n\nClick on the template preview to set coordinates.")
        self.coordinates_text.see(tk.END)
    
    def log_message(self, message):
        timestamp = pd.Timestamp.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {message}"
        self.log_text.insert(tk.END, formatted_message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
    def generate_cards(self):
        # Validate inputs
        required_fields = [
            (self.template_path.get(), "ID Card Template"),
            (self.photos_folder.get(), "Photos Folder"),
            (self.qr_folder.get(), "QR Codes Folder"),
            (self.excel_path.get(), "Excel File"),
            (self.output_folder.get(), "Output Folder")
        ]
        
        missing_fields = [field_name for field_value, field_name in required_fields if not field_value]
        
        if missing_fields:
            missing_text = "\n• ".join(missing_fields)
            messagebox.showerror("Missing Required Fields", 
                               f"Please select the following required fields:\n\n• {missing_text}")
            self.log_message(f"❌ Missing required fields: {', '.join(missing_fields)}")
            return
        
        if not self.coordinates:
            messagebox.showerror("Missing Coordinates", 
                               "Please set coordinates for all elements by clicking on the template preview or using manual input.")
            self.log_message("❌ No coordinates set")
            return
            
        try:
            self.log_message("🚀 Starting ID card generation...")
            
            # Create generator instance
            generator = IDCardGenerator(
                template_path=self.template_path.get(),
                photos_folder=self.photos_folder.get(),
                qr_folder=self.qr_folder.get(),
                excel_path=self.excel_path.get(),
                output_folder=self.output_folder.get(),
                coordinates=self.coordinates,
                log_callback=self.log_message,
                export_as_pdf_var=self.export_as_pdf,
                photo_frame_style=self.photo_frame_style.get(),
                font_color=self.font_color.get(),
                border_size=int(self.border_size.get()),
                border_color=self.border_color.get()
            )
            
            # Set custom font if provided
            if self.font_path.get():
                try:
                    # Pass font sizes dictionary to the generator
                    generator.set_font(self.font_path.get(), self.font_sizes)
                except Exception as e:
                    self.log_message(f"⚠️ Error setting font: {str(e)}")
            
            # Generate cards
            generator.generate_all_id_cards()
            
            success_msg = "🎉 ID cards generated successfully!"
            messagebox.showinfo("Success", success_msg)
            self.log_message(success_msg)
            
        except Exception as e:
            error_msg = f"An error occurred: {str(e)}"
            messagebox.showerror("Error", error_msg)
            self.log_message(f"❌ {error_msg}")

class IDCardGenerator:
    def __init__(self, template_path, photos_folder, qr_folder, excel_path, output_folder, coordinates=None, photo_size=(230, 230), qr_size=(120, 120), log_callback=None, export_as_pdf_var=None, photo_frame_style="circle", font_color="black", border_size=2, border_color="blue"):
        """
        Initialize the ID Card Generator.
        
        Args:
            template_path (str): Path to the ID card template PNG
            photos_folder (str): Path to folder containing student photos
            qr_folder (str): Path to folder containing QR code images
            excel_path (str): Path to Excel file with student details
            output_folder (str): Path to save generated ID cards
            coordinates (dict): Dictionary of coordinates for text and photo placement
            photo_size (tuple): Size of the photo (width, height)
            qr_size (tuple): Size of the QR code (width, height)
            log_callback (callable): Function to call for logging messages
            photo_frame_style (str): Style of photo frame ("circle" or "square")
            font_color (str): Color of the text
            border_size (int): Size of the photo border in pixels
            border_color (str): Color of the photo border
        """
        self.template_path = template_path
        self.photos_folder = photos_folder
        self.qr_folder = qr_folder
        self.excel_path = excel_path
        self.output_folder = output_folder
        self.log_callback = log_callback or (lambda x: None)
        self.export_as_pdf_var = export_as_pdf_var
        self.photo_frame_style = photo_frame_style
        self.font_color = font_color
        self.border_size = border_size
        self.border_color = border_color
        
        # Create output folder if it doesn't exist
        os.makedirs(output_folder, exist_ok=True)
        
        # Load template and get its dimensions
        self.template = Image.open(template_path)
        self.template_width, self.template_height = self.template.size
        
        # Use provided coordinates or defaults
        self.text_coordinates = {}
        self.photo_coordinates = coordinates.get('Photo', (293, 270))  # default
        self.qr_coordinates = coordinates.get('QR Code', (50, 50))  # default
        self.photo_size = photo_size
        self.qr_size = qr_size
        
        if coordinates:
            # Set photo and QR coordinates
            if 'Photo' in coordinates:
                self.photo_coordinates = coordinates['Photo']
            if 'QR Code' in coordinates:
                self.qr_coordinates = coordinates['QR Code']
            
            # Set text coordinates
            for key in ['Name', 'Class', 'Contact', 'Address', 'Guardian', 'Validity', 'Roll No.', 'RegNo']:
                if key in coordinates:
                    self.text_coordinates[key] = coordinates[key]

        # Mapping from card labels to Excel column names (case-insensitive matching will be used)
        # Map the label on the card to the likely column name in your Excel file.
        # Note: The keys in this dictionary are the labels shown on the card/GUI,
        # and the values are the corresponding column names in the Excel file.
        # For fields like Roll No. and RegNo, we provide a list of possible column names to handle variations.
        self.label_to_excel_column_map = {
            'Name': 'Name',
            'Class': 'Grade', # Map Class label on card to Grade column in Excel
            'Contact': 'PhoneNumber', # Map Contact label on card to PhoneNumber column in Excel
            'Address': 'Address',
            'Guardian': 'Guardian',
            'Validity': 'Validity',
            'Roll No.': ['Roll No.', 'Roll No', 'RollNo', 'Roll Number', 'RollNumber'], # Multiple possible column names for Roll No.
            'RegNo': ['RegNo', 'Reg No', 'Reg-No', 'Registration No', 'RegistrationNo'] # Added multiple possible column names for RegNo
        }
        
        # Initialize font variables
        self.font_path = None
        self.fonts = {}  # Dictionary to store fonts for different fields
        self.default_font = ImageFont.load_default()

    def set_font(self, font_path, font_sizes):
        """Set custom fonts for different text fields.

        Args:
            font_path (str): Path to the .ttf or .otf font file
            font_sizes (dict): Dictionary mapping field names to font sizes
        """
        try:
            self.font_path = font_path
            # Create font objects for each field with their respective sizes
            for field, size_var in font_sizes.items():
                try:
                    size = int(size_var.get())
                    self.fonts[field] = ImageFont.truetype(font_path, size)
                    self.log_callback(f"✅ Set font size {size} for {field}")
                except ValueError:
                    self.log_callback(f"⚠️ Invalid font size for {field}, using default size")
                    self.fonts[field] = ImageFont.truetype(font_path, 20)
        except Exception as e:
            self.log_callback(f"⚠️ Error loading font {font_path}: {str(e)}. Using default font.")
            self.fonts = {field: self.default_font for field in font_sizes.keys()}

    def get_font_for_field(self, field):
        """Get the appropriate font for a given field."""
        return self.fonts.get(field, self.default_font)

    def create_circular_mask(self, image):
        """Create a circular mask for the photo."""
        mask = Image.new('L', image.size, 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, image.size[0], image.size[1]), fill=255)
        return mask

    def process_photo(self, photo_path):
        """Process a photo by resizing it and applying the selected frame style."""
        try:
            # Open and resize the photo
            photo = Image.open(photo_path)
            photo = photo.resize(self.photo_size, Image.Resampling.LANCZOS)
            
            if self.photo_frame_style == "circle":
                # Create a circular mask with border
                mask = Image.new('L', self.photo_size, 0)
                draw = ImageDraw.Draw(mask)
                
                # Draw the main circle
                draw.ellipse((0, 0) + self.photo_size, fill=255)
                
                # Create border mask
                border_mask = Image.new('L', self.photo_size, 0)
                border_draw = ImageDraw.Draw(border_mask)
                
                # Draw border circle
                border_draw.ellipse((0, 0) + self.photo_size, outline=self.border_color, width=self.border_size)
                
                # Combine masks
                final_mask = Image.new('L', self.photo_size, 0)
                final_mask.paste(mask, (0, 0), mask)
                final_mask.paste(border_mask, (0, 0), border_mask)
                
                # Apply the mask
                output = Image.new('RGBA', self.photo_size, (0, 0, 0, 0))
                output.paste(photo, (0, 0))
                output.putalpha(final_mask)
                
                return output
            else:
                # For square photos, add a border
                if self.border_size > 0:
                    # Create a new image with border
                    bordered = Image.new('RGBA', self.photo_size, (0, 0, 0, 0))
                    # Create a drawing object for the border
                    border_draw = ImageDraw.Draw(bordered)
                    # Draw the rectangle border
                    border_draw.rectangle((0, 0) + self.photo_size, outline=self.border_color, width=self.border_size)
                    
                    # Calculate inner size
                    inner_size = (self.photo_size[0] - 2 * self.border_size, 
                                self.photo_size[1] - 2 * self.border_size)
                    # Resize photo to fit inside border
                    inner_photo = photo.resize(inner_size, Image.Resampling.LANCZOS)
                    # Paste photo in center
                    paste_x = (self.photo_size[0] - inner_size[0]) // 2
                    paste_y = (self.photo_size[1] - inner_size[1]) // 2
                    bordered.paste(inner_photo, (paste_x, paste_y))
                    return bordered
                return photo
        except Exception as e:
            self.log_callback(f"⚠️ Error processing photo {os.path.basename(photo_path)}: {str(e)}")
            return None

    def process_qr_code(self, qr_path):
        """Process a QR code image by resizing it."""
        try:
            # Open and resize the QR code
            qr_image = Image.open(qr_path)
            qr_image = qr_image.resize(self.qr_size, Image.Resampling.LANCZOS)
            return qr_image
        except Exception as e:
            self.log_callback(f"⚠️ Error processing QR code {os.path.basename(qr_path)}: {str(e)}")
            return None

    def generate_qr_code(self, student_data, ext_id_key):
        """Generate QR code for student data."""
        try:
            # Create QR code with student information
            qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                box_size=10,
                border=4,
            )
            # Add student data to QR code
            qr_data = {
                'ID': student_data.get(ext_id_key, 'Unknown'),
                'Name': student_data.get(self.label_to_excel_column_map.get('Name', 'Name'), ''),
                'Class': student_data.get(self.label_to_excel_column_map.get('Class', 'Grade'), '')
            }
            qr.add_data(str(qr_data))
            qr.make(fit=True)
            
            # Create QR code image
            qr_image = qr.make_image(fill_color="black", back_color="white")
            qr_image = qr_image.resize(self.qr_size, Image.Resampling.LANCZOS)
            return qr_image
        except Exception as e:
            self.log_callback(f"⚠️ Error generating QR code: {str(e)}")
            return None

    def generate_id_card(self, student_data, photo_path, ext_id_key):
        """Generate a single ID card."""
        try:
            student_id = student_data.get(ext_id_key, 'Unknown')
            self.log_callback(f"🔄 Processing student: {student_id}")
            
            # Create a copy of the template
            id_card = self.template.copy()
            
            photo_added = False
            # Process and paste the photo
            if photo_path and os.path.exists(photo_path):
                photo = self.process_photo(photo_path)
                if photo:
                    # Calculate center-aligned coordinates for photo
                    photo_x = self.photo_coordinates[0] - (self.photo_size[0] // 2)
                    photo_y = self.photo_coordinates[1] - (self.photo_size[1] // 2)
                    id_card.paste(photo, (photo_x, photo_y), photo) # Use photo with alpha channel for pasting
                    self.log_callback(f"  ✅ Photo added for {student_id}")
                    photo_added = True
            else:
                self.log_callback(f"  ⚠️ Photo not found for student {student_id}")
            
            qr_added = False
            # Look for and paste QR code
            ext_id = student_data.get(ext_id_key, 'Unknown')
            qr_path = None
            # Search for QR code file case-insensitively in the specified folder
            if os.path.exists(self.qr_folder):
                for filename in os.listdir(self.qr_folder):
                    if ext_id.lower() in filename.lower() and filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                        qr_path = os.path.join(self.qr_folder, filename)
                        break
            else:
                 self.log_callback(f"  ⚠️ QR Codes Folder not found: {self.qr_folder}")

            
            if qr_path and os.path.exists(qr_path):
                qr_image = self.process_qr_code(qr_path)
                if qr_image:
                    # Calculate center-aligned coordinates for QR code
                    qr_x = self.qr_coordinates[0] - (self.qr_size[0] // 2)
                    qr_y = self.qr_coordinates[1] - (self.qr_size[1] // 2)
                    id_card.paste(qr_image, (qr_x, qr_y)) # QR images are typically RGB/L, no mask needed
                    self.log_callback(f"  ✅ QR code added for {student_id}")
                    qr_added = True
            else:
                self.log_callback(f"  ⚠️ QR code not found for student {student_id}")
            
            # Add text information
            draw = ImageDraw.Draw(id_card)
            
            # Sort coordinates by y-axis to process text fields roughly top-to-bottom
            sorted_coords = sorted(self.text_coordinates.items(), key=lambda item: item[1][1])

            text_added_count = 0
            # Draw text for each defined coordinate label
            for field, (x, y) in sorted_coords:
                try:
                    # Check if a coordinate is set for this field label
                    if field not in self.text_coordinates:
                        # This case should ideally be caught before generation, but adding a check here for robustness
                        self.log_callback(f"  ⚠️ Coordinate not set for field '{field}', skipping text placement.")
                        continue
                        
                    # Find the actual column name in the Excel file corresponding to the card label
                    excel_column_key = self.label_to_excel_column_map.get(field)

                    # Handle cases where the mapping value is a list of possible column names
                    if isinstance(excel_column_key, list):
                        found_key = None
                        for possible_key in excel_column_key:
                            # Check if the possible column name exists in the student_data dictionary (from the Excel row)
                            if possible_key in student_data and pd.notna(student_data[possible_key]):
                                found_key = possible_key
                                break # Found a valid column name, no need to check others
                        actual_excel_key_in_dict = found_key
                    else:
                        # If the mapping value is not a list, use it directly as the column name
                        actual_excel_key_in_dict = excel_column_key
                    
                    # If the column name was found and exists in the student data
                    if actual_excel_key_in_dict and actual_excel_key_in_dict in student_data and pd.notna(student_data[actual_excel_key_in_dict]):
                        value = student_data[actual_excel_key_in_dict]
                        
                        # Special handling for Validity date to format it
                        if field == 'Validity':
                            try:
                                # Attempt to parse and format date, handle different input types
                                if isinstance(value, str):
                                    date_value = pd.to_datetime(value)
                                elif isinstance(value, pd.Timestamp):
                                     date_value = value
                                else:
                                    raise ValueError("Value is not a string or Timestamp") # Indicate failure for other types
                                text_data = date_value.strftime('%Y-%m-%d') # Format date as YYYY-MM-DD
                            except Exception as date_error:
                                # If date conversion/formatting fails, use the original value as string and log warning
                                text_data = str(value)
                                self.log_callback(f"  ⚠️ Could not format Validity date for {student_id}: {date_error}. Using original value.")
                        else:
                            # For other fields, convert the value to a string
                            text_data = str(value)
                        
                        # Get font color from instance variable, defaulting to black if not set or invalid
                        try:
                            font_color = self.font_color.get() if hasattr(self.font_color, 'get') else (self.font_color if self.font_color else 'black')
                            # Basic validation for color string (optional, but good practice)
                            if not isinstance(font_color, str) or not font_color: raise ValueError
                        except Exception as color_error:
                            font_color = 'black' # Fallback to black on error
                            self.log_callback(f"  ⚠️ Error getting font color: {color_error}. Using default black.")

                        # Get the appropriate font for this field
                        field_font = self.get_font_for_field(field)
                        
                        # Draw the text on the card with the field-specific font
                        draw.text((x, y), text_data, fill=font_color, font=field_font)
                        text_added_count += 1
                        self.log_callback(f"  ✅ Added text for '{field}': '{text_data}'")
                    else:
                        # Log if the column wasn't found in data or had no valid value
                        col_name = excel_column_key if isinstance(excel_column_key, str) else ", ".join(excel_column_key)
                        self.log_callback(f"  ⚠️ No data found for field '{field}' (looked for column(s): {col_name}), skipping text placement.")

                except Exception as e:
                    self.log_callback(f"  ❌ Error processing field '{field}' for student {student_id}: {str(e)}")
                    continue

            # Check if at least one element (text, photo, or QR) was successfully added to the card
            if text_added_count > 0 or photo_added or qr_added:
                self.log_callback(f"  ✅ Card generated for {student_id} with {text_added_count} text fields, Photo added: {photo_added}, QR added: {qr_added}")
                return id_card
            else:
                # If nothing could be added to the card, return None
                self.log_callback(f"  ❌ No data could be added to the card for student {student_id}. Skipping card generation.")
                return None

        except Exception as e:
            # Catch any critical errors during the processing of a single student's card
            student_id_for_log = student_data.get(ext_id_key, 'Unknown ID')
            self.log_callback(f"  ❌ Critical error generating ID card for {student_id_for_log}: {str(e)}")
            return None # Return None if a critical error occurred

    def generate_all_id_cards(self):
        """Generate ID cards for all students."""
        try:
            if not os.path.exists(self.excel_path):
                raise FileNotFoundError(f"Excel file not found: {self.excel_path}")
                
            self.log_callback(f"📖 Reading Excel file: {os.path.basename(self.excel_path)}")
            # Read student data from Excel
            df = pd.read_excel(self.excel_path)
            
            if df.empty:
                raise ValueError("Excel file is empty")
                
            total_students = len(df)
            self.log_callback(f"📊 Found {total_students} students in Excel file")
            
            # Log available columns for debugging
            self.log_callback(f"📋 Available columns: {', '.join(df.columns)}")
            
            # Log missing columns as warnings
            for field, required_cols in self.label_to_excel_column_map.items():
                if isinstance(required_cols, list):
                    if not any(col in df.columns for col in required_cols):
                        self.log_callback(f"⚠️ Optional field {field} not found (tried: {', '.join(required_cols)}) in Excel columns.")
                elif required_cols not in df.columns:
                    self.log_callback(f"⚠️ Optional field '{required_cols}' not found in Excel columns.")

            
            successful_cards = 0
            failed_cards = 0
            
            generated_images = [] # List to store generated images for PDF export

            try:
                # Process each student
                for index, row in df.iterrows():
                    self.log_callback(f"\n--- Processing Row {index + 1} ---")
                    student_data = row.to_dict()
                    self.log_callback(f"Raw row data: {student_data}")

                    # Find the actual dictionary key for 'ext_id' case-insensitively
                    ext_id_key_in_dict = None
                    for key in student_data.keys():
                        if key.lower() in ['ext_id', 'ext-id', 'extid', 'id']:
                            ext_id_key_in_dict = key
                            break

                    # If the key is not found in any case, or if the value is empty, skip the row
                    if not ext_id_key_in_dict or pd.isna(student_data.get(ext_id_key_in_dict)):
                        self.log_callback(f"⚠️ Row {index + 1}: 'EXT_ID' not found or is empty. Skipping row.")
                        failed_cards += 1
                        continue

                    # Get the ext_id value using the discovered key
                    ext_id = str(student_data[ext_id_key_in_dict]) # Ensure ext_id is a string for comparison
                    self.log_callback(f"Found EXT_ID: {ext_id} (using key '{ext_id_key_in_dict}')")

                    # Look for matching photo by checking if ext_id is in the filename
                    photo_path = None
                    # Search for photo file case-insensitively in the specified folder
                    if os.path.exists(self.photos_folder):
                         for filename in os.listdir(self.photos_folder):
                             # Check if ext_id (as string) is a substring of the lowercased filename
                             if ext_id.lower() in filename.lower() and filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                                 photo_path = os.path.join(self.photos_folder, filename)
                                 break # Found a match, no need to check other files
                    else:
                        self.log_callback(f"  ⚠️ Photos Folder not found: {self.photos_folder}. No photo search performed.")

                    if photo_path:
                        self.log_callback(f"Found photo file: {os.path.basename(photo_path)}")
                    else:
                        self.log_callback(f"  ⚠️ No photo found for EXT_ID: {ext_id} in {self.photos_folder}")

                    # Pass the student_data and the discovered ext_id_key to generate_id_card
                    generated_card_image = self.generate_id_card(student_data, photo_path, ext_id_key_in_dict)

                    if generated_card_image:
                        generated_images.append(generated_card_image)
                        self.log_callback(f"  ✅ Generated image for {ext_id} (added to PDF list)")
                        successful_cards += 1
                    else:
                        # If generate_id_card returned None (due to error or no data added)
                        self.log_callback(f"  ❌ Failed to generate card for {ext_id} (generate_id_card returned None)")
                        failed_cards += 1

            except KeyboardInterrupt:
                self.log_callback("\n⚠️ ID card generation was interrupted by user.")
                messagebox.showwarning("Generation Interrupted", 
                                     f"ID card generation was interrupted.\n\nProgress:\n• {successful_cards} cards generated\n• {failed_cards} failed")
                # Continue to final summary and PDF saving based on images collected so far
                pass # Allow execution to continue to the final summary and PDF save

            # Final summary
            self.log_callback(f"\n🎯 Generation Summary:")
            self.log_callback(f"  • Total students processed: {total_students}")
            self.log_callback(f"  • Successful cards: {successful_cards}")
            self.log_callback(f"  • Failed cards: {failed_cards}")

            # --- PDF Export Logic ---
            if generated_images and self.export_as_pdf_var.get():
                self.log_callback("✅ 'Export as single PDF (A4 Landscape)' is checked. Preparing PDF...")
                # Ensure output folder exists
                os.makedirs(self.output_folder, exist_ok=True)

                pdf_output_path = os.path.join(self.output_folder, "all_id_cards.pdf")
                self.log_callback(f"📦 Arranging {len(generated_images)} cards onto A4 pages for PDF export...")

                a4_images = [] # List to store the final A4 pages

                # Define A4 dimensions in pixels at 300 DPI (approx 3508x2480 portrait, so 2480x3508 landscape)
                # Using common A4 aspect ratio scaled for reasonable pixel dimensions
                a4_width_px = 3508 # Landscape width at 300 DPI (approx 11.69 inches * 300 DPI)
                a4_height_px = 2480 # Landscape height at 300 DPI (approx 8.27 inches * 300 DPI)
                a4_size = (a4_width_px, a4_height_px)

                self.log_callback(f"📏 A4 Landscape size (px): {a4_size}")

                # Define grid layout (5 columns, 2 rows) based on user's example image
                num_cols = 5
                num_rows = 2
                cards_per_page = num_cols * num_rows

                # Calculate space for each card, including padding
                card_slot_width = a4_width_px // num_cols
                card_slot_height = a4_height_px // num_rows

                self.log_callback(f"📐 Card slot size (px): ({card_slot_width}, {card_slot_height}) for {num_cols}x{num_rows} grid")
                
                # Define padding around each card within its slot
                card_padding_px = 30 # Adjust this value to control the gap
                self.log_callback(f"Padding around each card (px): {card_padding_px}")

                # Process images in batches of cards_per_page
                for i in range(0, len(generated_images), cards_per_page):
                    self.log_callback(f"📄 Creating A4 page {len(a4_images) + 1} for batch starting at index {i}")
                    a4_page = Image.new('RGB', a4_size, (255, 255, 255)) # Create a blank white A4 landscape page
                    draw = ImageDraw.Draw(a4_page)

                    # Optional: Draw faint guide lines for debugging layout
                    # for col in range(num_cols + 1): draw.line([(col*card_slot_width, 0), (col*card_slot_width, a4_height_px)], fill=(200,200,200), width=1)
                    # for row in range(num_rows + 1): draw.line([(0, row*card_slot_height), (a4_width_px, row*card_slot_height)], fill=(200,200,200), width=1)

                    batch = generated_images[i : i + cards_per_page]

                    for j, card_img in enumerate(batch):
                        col_index = j % num_cols
                        row_index = j // num_cols

                        # Calculate the top-left position of the current slot
                        slot_x = col_index * card_slot_width
                        slot_y = row_index * card_slot_height
                        
                        self.log_callback(f"    🖼️ Processing card {j+1} in batch, slot ({col_index}, {row_index}) at position ({slot_x}, {slot_y})")

                        # Resize card image to fit within the slot while maintaining aspect ratio
                        # Calculate the maximum possible dimensions for the card while fitting within the slot, considering padding
                        max_card_width = card_slot_width - 2 * card_padding_px
                        max_card_height = card_slot_height - 2 * card_padding_px
                        
                        self.log_callback(f"    📦 Max card size within padding (px): ({max_card_width}, {max_card_height})")

                        # Get original card image size
                        original_card_width, original_card_height = card_img.size
                        self.log_callback(f"    📏 Original card size (px): ({original_card_width}, {original_card_height})")

                        # Calculate scaling factor to fit within the slot
                        width_scale = max_card_width / original_card_width
                        height_scale = max_card_height / original_card_height

                        # Use the minimum scale factor to maintain aspect ratio
                        scale_factor = min(width_scale, height_scale)
                        self.log_callback(f"    🔎 Scaling factor: {scale_factor:.4f}")

                        # Calculate the new size for the card image
                        new_card_width = int(original_card_width * scale_factor)
                        new_card_height = int(original_card_height * scale_factor)
                        self.log_callback(f"    ✨ Resized card size (px): ({new_card_width}, {new_card_height})")

                        # Resize the card image using LANCZOS for quality
                        resized_card = card_img.resize((new_card_width, new_card_height), Image.Resampling.LANCZOS)

                        # Calculate the paste position to center the resized card within its slot
                        paste_x = slot_x + (card_slot_width - new_card_width) // 2
                        paste_y = slot_y + (card_slot_height - new_card_height) // 2
                        
                        self.log_callback(f"    📍 Pasting resized card at (px): ({paste_x}, {paste_y})")

                        # Paste the resized card image onto the A4 page
                        # Check if the resized_card has an alpha channel before pasting with mask
                        if resized_card.mode == 'RGBA':
                             a4_page.paste(resized_card, (paste_x, paste_y), resized_card)
                        else:
                             a4_page.paste(resized_card, (paste_x, paste_y))

                    a4_images.append(a4_page)
                    self.log_callback(f"  ✅ Finished creating A4 page {len(a4_images)}")

                try:
                    if a4_images:
                        # Convert A4 page images to RGB mode before saving as PDF
                        rgb_a4_images = [img.convert('RGB') for img in a4_images]

                        # Save the first A4 image, then append the rest
                        rgb_a4_images[0].save(pdf_output_path, save_all=True, append_images=rgb_a4_images[1:], format='PDF')
                        self.log_callback(f"🎉 Successfully saved A4 PDF to: {os.path.basename(pdf_output_path)}")
                    else:
                        self.log_callback("⚠️ No A4 pages were generated to save as PDF.")

                except Exception as e:
                    self.log_callback(f"❌ Error saving PDF: {str(e)}")
                    messagebox.showerror("PDF Save Error", f"Error saving PDF: {str(e)}")

            # Log if PDF export was skipped because the option was not selected or no images were generated
            elif not self.export_as_pdf_var.get():
                 self.log_callback("⏭️ PDF export option not selected. Skipping PDF generation.")
            elif not generated_images:
                 self.log_callback("⚠️ No cards generated. Skipping PDF save.")

        except FileNotFoundError as e:
            self.log_callback(f"❌ File not found error: {str(e)}")
            messagebox.showerror("File Not Found Error", f"Error: {str(e)}")
        except ValueError as e:
             self.log_callback(f"❌ Data error: {str(e)}")
             messagebox.showerror("Data Error", f"Error: {str(e)}")
        except Exception as e:
            self.log_callback(f"❌ Critical error during bulk generation: {str(e)}")
            messagebox.showerror("Generation Error", f"Critical error during generation: {str(e)}")

def main():
    root = tk.Tk()
    app = IDCardGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()