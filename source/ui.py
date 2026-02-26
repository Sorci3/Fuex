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
        
        # --- Palette de couleurs (Mode Sombre Moderne) ---
        self.colors = {
            "bg_main": "#1e1e2e",       # Arrière-plan principal (Dark blue/gray)
            "bg_card": "#2a2a3e",       # Arrière-plan des panneaux et sections
            "bg_drop": "#313244",       # Arrière-plan des zones dynamiques (glisser-déposer)
            "fg_text": "#cdd6f4",       # Texte principal
            "fg_title": "#89b4fa",      # Accent principal (Bleu)
            "fg_accent1": "#f38ba8",    # Accent pour Main (Rouge clair / Rose)
            "fg_accent2": "#a6e3a1",    # Accent pour Source (Vert clair)
            "btn_bg": "#89b4fa",        # Fond des boutons normaux
            "btn_fg": "#1e1e2e",        # Texte des boutons (fort contraste)
            "btn_hover": "#b4befe",     # Hover des boutons
            "btn_success": "#a6e3a1",   # Bouton fusion
            "btn_success_hover": "#94e2d5", 
            "border": "#45475a"         # Bordures
        }
        
        self.root.configure(bg=self.colors["bg_main"])

        # Paramétrage du style
        self.style = ttk.Style()
        if 'clam' in self.style.theme_names():
            self.style.theme_use('clam')
            
        self.setup_styles()
        
        self.main_path = tk.StringVar()
        self.source_path = tk.StringVar()
        self.mapping_widgets = {}

        self.create_widgets()

    def setup_styles(self):
        # Configuration globale ttk
        self.style.configure(".", background=self.colors["bg_main"], foreground=self.colors["fg_text"], font=("Helvetica", 10))
        
        # LabelFrames
        self.style.configure("Card.TLabelframe", background=self.colors["bg_card"], bordercolor=self.colors["border"], borderwidth=1, relief="solid")
        self.style.configure("Card.TLabelframe.Label", background=self.colors["bg_card"], foreground=self.colors["fg_title"], font=("Helvetica", 11, "bold"))

        # Boutons standards
        self.style.configure("TButton", background=self.colors["btn_bg"], foreground=self.colors["btn_fg"], borderwidth=0, focuscolor=self.colors["btn_bg"], font=("Helvetica", 10, "bold"), padding=6)
        self.style.map("TButton", background=[("active", self.colors["btn_hover"])])
        
        # Bouton de Fusion
        self.style.configure("Success.TButton", background=self.colors["btn_success"], foreground=self.colors["btn_fg"], borderwidth=0, font=("Helvetica", 12, "bold"), padding=12)
        self.style.map("Success.TButton", 
                       background=[("active", self.colors["btn_success_hover"]), ("disabled", self.colors["bg_drop"])], 
                       foreground=[("disabled", "#6c7086")])

        # Scrollbar
        self.style.configure("Vertical.TScrollbar", background=self.colors["bg_card"], troughcolor=self.colors["bg_main"], bordercolor=self.colors["bg_main"], arrowcolor=self.colors["fg_text"])
        
        # Combobox
        self.style.configure("TCombobox", fieldbackground=self.colors["bg_drop"], background=self.colors["bg_card"], foreground=self.colors["fg_text"], arrowcolor=self.colors["fg_text"], bordercolor=self.colors["border"])
        self.style.map("TCombobox", 
                       fieldbackground=[('readonly', self.colors["bg_drop"])], 
                       selectbackground=[('readonly', self.colors["btn_bg"])], 
                       selectforeground=[('readonly', self.colors["btn_fg"])])
                       
        # Fix pour le menu déroulant sur Windows
        self.root.option_add('*TCombobox*Listbox.background', self.colors["bg_drop"])
        self.root.option_add('*TCombobox*Listbox.foreground', self.colors["fg_text"])
        self.root.option_add('*TCombobox*Listbox.selectBackground', self.colors["btn_bg"])
        self.root.option_add('*TCombobox*Listbox.selectForeground', self.colors["btn_fg"])
        self.root.option_add('*TCombobox*Listbox.font', ("Helvetica", 10))


    def create_widgets(self):
        # --- Section 1 : Zones de Drag & Drop ---
        frame_top = tk.Frame(self.root, bg=self.colors["bg_main"])
        frame_top.pack(fill="x", padx=20, pady=20)

        # Zone Fichier Main
        self.frame_main = ttk.LabelFrame(frame_top, text=" 1. Excel Principal (Main) ", style="Card.TLabelframe")
        self.frame_main.pack(side="left", fill="both", expand=True, padx=(0, 10), ipadx=5, ipady=5)
        
        self.lbl_main = tk.Label(self.frame_main, text="Glissez l'Excel Main ici\n\nou", bg=self.colors["bg_drop"], fg=self.colors["fg_text"], font=("Helvetica", 11), width=30, height=4, relief="flat", cursor="hand2")
        self.lbl_main.pack(pady=(15, 5), padx=15, fill="x")
        ttk.Button(self.frame_main, text="Parcourir...", command=self.load_main).pack(pady=(5, 10))
        tk.Label(self.frame_main, textvariable=self.main_path, bg=self.colors["bg_card"], fg=self.colors["fg_accent1"], font=("Helvetica", 9, "bold"), wraplength=300).pack(pady=5, padx=10)

        # Zone Fichier Source
        self.frame_source = ttk.LabelFrame(frame_top, text=" 2. Excel à ajouter (Source) ", style="Card.TLabelframe")
        self.frame_source.pack(side="right", fill="both", expand=True, padx=(10, 0), ipadx=5, ipady=5)
        
        self.lbl_source = tk.Label(self.frame_source, text="Glissez l'Excel Source ici\n\nou", bg=self.colors["bg_drop"], fg=self.colors["fg_text"], font=("Helvetica", 11), width=30, height=4, relief="flat", cursor="hand2")
        self.lbl_source.pack(pady=(15, 5), padx=15, fill="x")
        ttk.Button(self.frame_source, text="Parcourir...", command=self.load_source).pack(pady=(5, 10))
        tk.Label(self.frame_source, textvariable=self.source_path, bg=self.colors["bg_card"], fg=self.colors["fg_accent2"], font=("Helvetica", 9, "bold"), wraplength=300).pack(pady=5, padx=10)

        # Actions clic additionnelles sur les labels
        self.lbl_main.bind("<Button-1>", lambda e: self.load_main())
        self.lbl_source.bind("<Button-1>", lambda e: self.load_source())

        # Configuration du Drag & Drop
        self.frame_main.drop_target_register(DND_FILES)
        self.frame_main.dnd_bind('<<Drop>>', self.drop_main)
        self.lbl_main.drop_target_register(DND_FILES)
        self.lbl_main.dnd_bind('<<Drop>>', self.drop_main)

        self.frame_source.drop_target_register(DND_FILES)
        self.frame_source.dnd_bind('<<Drop>>', self.drop_source)
        self.lbl_source.drop_target_register(DND_FILES)
        self.lbl_source.dnd_bind('<<Drop>>', self.drop_source)


        # --- Section 2 : Mapping des colonnes ---
        self.frame_mapping = ttk.LabelFrame(self.root, text=" 3. Mapping des colonnes (Main - Source) ", style="Card.TLabelframe")
        self.frame_mapping.pack(fill="both", expand=True, padx=20, pady=5)
        
        self.lbl_mapping_empty = tk.Label(self.frame_mapping, text="Sélectionnez d'abord les deux fichiers Excel ci-dessus.", bg=self.colors["bg_card"], fg="#6c7086", font=("Helvetica", 11, "italic"))
        self.lbl_mapping_empty.pack(expand=True)

        # --- Section 3 : Fusion ---
        frame_action = tk.Frame(self.root, bg=self.colors["bg_main"])
        frame_action.pack(fill="x", padx=20, pady=15)
        
        self.btn_merge = ttk.Button(frame_action, text="FUSIONNER LES EXCEL", style="Success.TButton", state="disabled", command=self.process_merge)
        # On utilise padding via grid ou pack interne
        self.btn_merge.pack(pady=10, fill="x")

    # --- Utilitaires Drag & Drop ---
    def clean_path(self, path):
        """Nettoie le chemin du fichier (enlève les accolades ajoutées par tkdnd sur Windows)"""
        if path.startswith('{') and path.endswith('}'):
            return path[1:-1]
        return path

    def drop_main(self, event):
        filepath = self.clean_path(event.data)
        if filepath.lower().endswith(('.xlsx', '.xls')):
            self.lbl_main.configure(bg=self.colors["bg_card"], text="✅ Fichier Main chargé\nDrop à nouveau pour changer")
            self.main_path.set(filepath)
            self.check_ready_for_mapping()
        else:
            messagebox.showwarning("Erreur", "Veuillez glisser un fichier Excel (.xlsx ou .xls)")

    def drop_source(self, event):
        filepath = self.clean_path(event.data)
        if filepath.lower().endswith(('.xlsx', '.xls')):
            self.lbl_source.configure(bg=self.colors["bg_card"], text="✅ Fichier Source chargé\nDrop à nouveau pour changer")
            self.source_path.set(filepath)
            self.check_ready_for_mapping()
        else:
            messagebox.showwarning("Erreur", "Veuillez glisser un fichier Excel (.xlsx ou .xls)")

    # --- Boutons Parcourir (Fallback) ---
    def load_main(self):
        filepath = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
        if filepath:
            self.lbl_main.configure(bg=self.colors["bg_card"], text="✅ Fichier Main chargé\nCliquez pour changer")
            self.main_path.set(filepath)
            self.check_ready_for_mapping()

    def load_source(self):
        filepath = filedialog.askopenfilename(filetypes=[("Fichiers Excel", "*.xlsx *.xls")])
        if filepath:
            self.lbl_source.configure(bg=self.colors["bg_card"], text="✅ Fichier Source chargé\nCliquez pour changer")
            self.source_path.set(filepath)
            self.check_ready_for_mapping()

    # --- Logique d'interface ---
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
            
            options = ["Ignorer"] + source_cols

            # Canvas et Scrollbar au cas où il y a beaucoup de colonnes
            canvas = tk.Canvas(self.frame_mapping, borderwidth=0, highlightthickness=0, bg=self.colors["bg_card"])
            scrollbar = ttk.Scrollbar(self.frame_mapping, orient="vertical", command=canvas.yview)
            scrollable_frame = tk.Frame(canvas, bg=self.colors["bg_card"])

            scrollable_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )
            
            # Ajustement du fond du canvas avec bind
            canvas.bind("<Configure>", lambda e: canvas.itemconfig('frame', width=canvas.winfo_width()))

            canvas.create_window((0, 0), window=scrollable_frame, anchor="nw", tags="frame")
            canvas.configure(yscrollcommand=scrollbar.set)

            canvas.pack(side="left", fill="both", expand=True, padx=5, pady=5)
            scrollbar.pack(side="right", fill="y")

            # Entêtes
            tk.Label(scrollable_frame, text="Colonne de destination (Main)", font=("Helvetica", 10, "bold", "underline"), bg=self.colors["bg_card"], fg=self.colors["fg_accent1"]).grid(row=0, column=0, sticky="e", pady=(10, 15), padx=10)
            tk.Label(scrollable_frame, text="Valeur à importer (Source)", font=("Helvetica", 10, "bold", "underline"), bg=self.colors["bg_card"], fg=self.colors["fg_accent2"]).grid(row=0, column=1, sticky="w", pady=(10, 15), padx=10)

            for i, main_col in enumerate(main_cols, start=1):
                tk.Label(scrollable_frame, text=f"{main_col}", font=("Helvetica", 10, "bold"), bg=self.colors["bg_card"], fg=self.colors["fg_text"]).grid(row=i, column=0, sticky="e", pady=6, padx=10)
                
                cb = ttk.Combobox(scrollable_frame, values=options, state="readonly", width=40)
                if main_col in source_cols:
                    cb.set(main_col)
                else:
                    cb.set("Ignorer")
                    
                cb.grid(row=i, column=1, sticky="w", pady=6, padx=10)
                self.mapping_widgets[main_col] = cb

        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire les colonnes.\n{e}")

    def process_merge(self):
        mapping_dict = {}
        for main_col, combobox in self.mapping_widgets.items():
            mapping_dict[main_col] = combobox.get()

        output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile="Excel_Fusionne.xlsx")
        
        if not output_path:
            return 

        try:
            logic.merge_files(self.main_path.get(), self.source_path.get(), mapping_dict, output_path)
            messagebox.showinfo("Succès", "Les fichiers ont été fusionnés avec succès !")
        except Exception as e:
            messagebox.showerror("Erreur", f"Une erreur s'est produite lors de la fusion :\n{e}")
