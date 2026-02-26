import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from tkinterdnd2 import DND_FILES
import os
import source.logic as logic

class ExcelMergerUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Fuex")
        self.root.geometry("800x650")
        
        self.colors = {
            "bg_main": "#1e1e2e",       
            "bg_card": "#2a2a3e",       
            "bg_drop": "#313244",       
            "fg_text": "#cdd6f4",       
            "fg_title": "#89b4fa",      
            "fg_accent1": "#f38ba8",    
            "fg_accent2": "#a6e3a1",    
            "btn_bg": "#89b4fa",        
            "btn_fg": "#1e1e2e",        
            "btn_hover": "#b4befe",     
            "btn_success": "#a6e3a1",   
            "btn_success_hover": "#94e2d5", 
            "border": "#45475a"         
        }
        
        self.root.configure(bg=self.colors["bg_main"])

        self.style = ttk.Style()
        if 'clam' in self.style.theme_names():
            self.style.theme_use('clam')
            
        self.setup_styles()
        
        self.main_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.mapping_widgets = {}

        self.create_widgets()

    def setup_styles(self):
        self.style.configure(".", background=self.colors["bg_main"], foreground=self.colors["fg_text"], font=("Helvetica", 10))
        
        self.style.configure("Card.TLabelframe", background=self.colors["bg_card"], bordercolor=self.colors["border"], borderwidth=1, relief="solid")
        self.style.configure("Card.TLabelframe.Label", background=self.colors["bg_card"], foreground=self.colors["fg_title"], font=("Helvetica", 11, "bold"))

        self.style.configure("TButton", background=self.colors["btn_bg"], foreground=self.colors["btn_fg"], borderwidth=0, focuscolor=self.colors["btn_bg"], font=("Helvetica", 10, "bold"), padding=6)
        self.style.map("TButton", background=[("active", self.colors["btn_hover"])])
        
        self.style.configure("Success.TButton", background=self.colors["btn_success"], foreground=self.colors["btn_fg"], borderwidth=0, font=("Helvetica", 12, "bold"), padding=12)
        self.style.map("Success.TButton", 
                       background=[("active", self.colors["btn_success_hover"]), ("disabled", self.colors["bg_drop"])], 
                       foreground=[("disabled", "#6c7086")])

        self.style.configure("Vertical.TScrollbar", background=self.colors["bg_card"], troughcolor=self.colors["bg_main"], bordercolor=self.colors["bg_main"], arrowcolor=self.colors["fg_text"])
        
        self.style.configure("TCombobox", fieldbackground=self.colors["bg_drop"], background=self.colors["bg_card"], foreground=self.colors["fg_text"], arrowcolor=self.colors["fg_text"], bordercolor=self.colors["border"])
        self.style.map("TCombobox", 
                       fieldbackground=[('readonly', self.colors["bg_drop"])], 
                       selectbackground=[('readonly', self.colors["btn_bg"])], 
                       selectforeground=[('readonly', self.colors["btn_fg"])])
                       
        self.root.option_add('*TCombobox*Listbox.background', self.colors["bg_drop"])
        self.root.option_add('*TCombobox*Listbox.foreground', self.colors["fg_text"])
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.colors["btn_bg"])
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.colors["btn_fg"])
        self.root.option_add('*TCombobox*Listbox.font', ("Helvetica", 10))


    def create_widgets(self):
        # Drag and drop
        frame_top = tk.Frame(self.root, bg=self.colors["bg_main"])
        frame_top.pack(fill="x", padx=20, pady=20)

        self.frame_main = ttk.LabelFrame(frame_top, text=" 1. Main Excel ", style="Card.TLabelframe")
        self.frame_main.pack(side="left", fill="both", expand=True, padx=(0, 10), ipadx=5, ipady=5)
        
        self.lbl_main = tk.Label(self.frame_main, text="Drag & Drop Main Excel here\n\nor", bg=self.colors["bg_drop"], fg=self.colors["fg_text"], font=("Helvetica", 11), width=30, height=4, relief="flat", cursor="hand2")
        self.lbl_main.pack(pady=(15, 5), padx=15, fill="x")
        ttk.Button(self.frame_main, text="Browse...", command=self.load_main).pack(pady=(5, 10))
        tk.Label(self.frame_main, textvariable=self.main_path, bg=self.colors["bg_card"], fg=self.colors["fg_accent1"], font=("Helvetica", 9, "bold"), wraplength=300).pack(pady=5, padx=10)

        self.frame_source = ttk.LabelFrame(frame_top, text=" 2. Excel to add (Source) ", style="Card.TLabelframe")
        self.frame_source.pack(side="right", fill="both", expand=True, padx=(10, 0), ipadx=5, ipady=5)
        
        self.lbl_source = tk.Label(self.frame_source, text="Drag & Drop Source Excel here\n\nor", bg=self.colors["bg_drop"], fg=self.colors["fg_text"], font=("Helvetica", 11), width=30, height=4, relief="flat", cursor="hand2")
        self.lbl_source.pack(pady=(15, 5), padx=15, fill="x")
        ttk.Button(self.frame_source, text="Browse...", command=self.load_source).pack(pady=(5, 10))
        tk.Label(self.frame_source, textvariable=self.source_path, bg=self.colors["bg_card"], fg=self.colors["fg_accent2"], font=("Helvetica", 9, "bold"), wraplength=300).pack(pady=5, padx=10)

        self.lbl_main.bind("<Button-1>", lambda e: self.load_main())
        self.lbl_source.bind("<Button-1>", lambda e: self.load_source())

        self.frame_main.drop_target_register(DND_FILES)
        self.frame_main.dnd_bind('<<Drop>>', self.drop_main)
        self.lbl_main.drop_target_register(DND_FILES)
        self.lbl_main.dnd_bind('<<Drop>>', self.drop_main)

        self.frame_source.drop_target_register(DND_FILES)
        self.frame_source.dnd_bind('<<Drop>>', self.drop_source)
        self.lbl_source.drop_target_register(DND_FILES)
        self.lbl_source.dnd_bind('<<Drop>>', self.drop_source)


        # Mapping
        self.frame_mapping = ttk.LabelFrame(self.root, text=" 3. Column Mapping (Main - Source) ", style="Card.TLabelframe")
        self.frame_mapping.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.lbl_mapping_empty = tk.Label(self.frame_mapping, text="Please select both Excel files above first.", bg=self.colors["bg_card"], fg="#6c7086", font=("Helvetica", 11, "italic"))
        self.lbl_mapping_empty.pack(expand=True)

        # Merge
        frame_action = tk.Frame(self.root, bg=self.colors["bg_main"])
        frame_action.pack(fill="x", padx=20, pady=15)
        
        self.btn_merge = ttk.Button(frame_action, text="MERGE EXCEL FILES", style="Success.TButton", state="disabled", command=self.process_merge)
        self.btn_merge.pack(pady=10, fill="x")

    def clean_path(self, path):
        """Clean the file path (removes braces added by tkdnd on Windows)"""
        if path.startswith('{') and path.endswith('}'):
            return path[1:-1]
        return path

    def drop_main(self, event):
        filepath = self.clean_path(event.data)
        if filepath.lower().endswith(('.xlsx', '.xls')):
            self.lbl_main.configure(bg=self.colors["bg_card"], text="Main File loaded\nDrop again to change")
            self.main_path.set(filepath)
            self.check_ready_for_mapping()
        else:
            messagebox.showwarning("Error", "Please drop an Excel file (.xlsx or .xls)")

    def drop_source(self, event):
        filepath = self.clean_path(event.data)
        if filepath.lower().endswith(('.xlsx', '.xls')):
            self.lbl_source.configure(bg=self.colors["bg_card"], text="Source File loaded\nDrop again to change")
            self.source_path.set(filepath)
            self.check_ready_for_mapping()
        else:
            messagebox.showwarning("Error", "Please drop an Excel file (.xlsx or .xls)")

    def load_main(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filepath:
            self.lbl_main.configure(bg=self.colors["bg_card"], text="Main File loaded\nClick to change")
            self.main_path.set(filepath)
            self.check_ready_for_mapping()

    def load_source(self):
        filepath = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xls")])
        if filepath:
            self.lbl_source.configure(bg=self.colors["bg_card"], text="Source File loaded\nClick to change")
            self.source_path.set(filepath)
            self.check_ready_for_mapping()

    def check_ready_for_mapping(self):
        if self.main_path.get() and self.source_path.get():
            self.generate_mapping_ui()
            self.btn_merge.config(state="normal")

    def generate_mapping_ui(self):
        for widget in self.frame_mapping.winfo_children():
            widget.destroy()
        self.mapping_widgets.clear()

        try:
            main_cols = logic.get_excel_columns(self.main_path.get())
            source_cols = logic.get_excel_columns(self.source_path.get())
            
            options = ["Ignore"] + source_cols

            canvas = tk.Canvas(self.frame_mapping, borderwidth=0, highlightthickness=0, bg=self.colors["bg_card"])
            scrollbar = ttk.Scrollbar(self.frame_mapping, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=self.colors["bg_card"])

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            canvas.bind("<Configure>", lambda e: canvas.itemconfig('frame', width=canvas.winfo_width()))

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags="frame")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            scrollbar.pack(side="right", fill="y")

            tk.Label(scrollable_frame, text="Destination Column (Main)", font=("Helvetica", 10, "bold", "underline"), bg=self.colors["bg_card"], fg=self.colors["fg_accent1"]).grid(row=0, column=0, sticky="e", pady=(10, 15), padx=10)
            tk.Label(scrollable_frame, text="Value to Import (Source)", font=("Helvetica", 10, "bold", "underline"), bg=self.colors["bg_card"], fg=self.colors["fg_accent2"]).grid(row=0, column=1, sticky="w", pady=(10, 15), padx=10)

            for i, main_col in enumerate(main_cols, start=1):
                tk.Label(scrollable_frame, text=f"{main_col}", font=("Helvetica", 10, "bold"), bg=self.colors["bg_card"], fg=self.colors["fg_text"]).grid(row=i, column=0, sticky="e", pady=6, padx=10)
                
                cb = ttk.Combobox(scrollable_frame, values=options, state="readonly", width=40)
                if main_col in source_cols:
                    cb.set(main_col)
                else:
                    cb.set("Ignore")
                    
                cb.grid(row=i, column=1, sticky="w", pady=6, padx=10)
                self.mapping_widgets[main_col] = cb

        except Exception as e:
            messagebox.showerror("Error", f"Could not read columns.\n{e}")

    def process_merge(self):
        mapping_dict = {}
        for main_col, combobox in self.mapping_widgets.items():
            mapping_dict[main_col] = combobox.get()

        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="Merged_Excel.xlsx")
        
        if not output_path:
            return 

        try:
            logic.merge_files(self.main_path.get(), self.source_path.get(), mapping_dict, output_path)
            messagebox.showinfo("Success", "The files were merged successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during merging:\n{e}")
