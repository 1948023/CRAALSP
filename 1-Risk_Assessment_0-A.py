import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import math
import os
from datetime import datetime

# Import for Word export/import
try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

class RiskAssessmentTool:
    """Optimized Risk Assessment Tool for space missions"""
      # Color configuration
    COLORS = {
        'primary': '#4a90c2', 'success': '#28a745', 'white': '#ffffff',
        'light': '#f8f9fa', 'dark': '#2c3e50', 'gray': '#6c757d',
        'criteria_header': '#5a67d8', 'criteria_bg': '#edf2f7',
        'asset_header': '#38b2ac', 'asset_bg': '#f0fff4'
    }    
    # Main table data
    THREATS = [
        "Data Corruption", "Physical/Logical Attack", "Interception/Eavesdropping",
        "Jamming", "Denial-of-Service", "Masquerade/Spoofing", "Replay",
        "Software Threats", "Unauthorized Access/Hijacking", 
        "Tainted hardware components", "Supply Chain"    ]
    # Standard asset categories
    ASSET_CATEGORIES = [
        ("Ground", "Ground Stations"), ("Ground", "Mission Control"),
        ("Ground", "Data Processing Centers"), ("Ground", "Remote Terminals"),
        ("Ground", "User Ground Segment"), ("Space", "Platform"),
        ("Space", "Payload"), ("Link", "Link"), ("User", "User")    ]
    
    # Criteria table data (5x6 + header)
    CRITERIA_DATA = [
        ["Criteria", "Score 1 (Very Low)", "Score 2 (Low)", "Score 3 (Moderate)", "Score 4 (High)", "Score 5 (Very High)"],
        ["Vulnerability Level", "No know or already resolved vulnerabilities", "Know vulnerability, mitigate throught hardening and patches", "Know vulnerability, but only partially mitigated", "Known vulnerability, with no effective mitigation", "Actively exploitable vulnerability, with no defense"],
        ["Access Control", "Access strongly protected by physical/logical measures and isolated environment", "Moderately protected access with some isolation controls", "Standard access protection with basic controls", "Access easily accessible by remote attackers", "Completely open or physically accessible access"],
        ["Defense Capability", "Multi-level validated countermeasures with real-time automated detection", "Robust countermeasures with automated but decentralized detection", "Limited countermeasures with manual detection only", "Weak countermeasure with occasional detection", "No countermeasures or detection capabilities"],
        ["Operational Impact", "No impact thanks to redundancy with predefined automated response", "Negligible impact, quick response and system easily restored", "Medium impact with manual response, but mission continues", "Serious impact with slow response, mission temporarily interrupted", "Permanent loss of assets or mission with no response capability"],
        ["Recovery Time", "Immediate restoration with automated procedures", "Quick recovery within hours to days using standard procedures", "Manual recovery requiring weeks of coordinated effort", "Complex recovery requiring months of specialized intervention", "Impossible recovery or permanent system loss"]    ]
    
    # Risk matrix
    RISK_MATRIX = {
        ("Very High", "Very High"): "Very High", ("Very High", "High"): "Very High",
        ("Very High", "Medium"): "High", ("Very High", "Low"): "High",
        ("Very High", "Very Low"): "Medium", ("High", "Very High"): "Very High",
        ("High", "High"): "High", ("High", "Medium"): "High",
        ("High", "Low"): "Medium", ("High", "Very Low"): "Low",
        ("Medium", "Very High"): "High", ("Medium", "High"): "High",
        ("Medium", "Medium"): "Medium", ("Medium", "Low"): "Low",
        ("Medium", "Very Low"): "Low", ("Low", "Very High"): "Medium",
        ("Low", "High"): "Medium", ("Low", "Medium"): "Low",
        ("Low", "Low"): "Low", ("Low", "Very Low"): "Very Low",
        ("Very Low", "Very High"): "Low", ("Very Low", "High"): "Low",
        ("Very Low", "Medium"): "Low", ("Very Low", "Low"): "Very Low",
        ("Very Low", "Very Low"): "Very Low"    }
    # Security controls for each threat
    THREAT_COUNTERMEASURES = {
        "Data Corruption": [
            "Configuration Management", "Tamper resistant body", "Tamper Protection", 
            "Disable Physical Ports", "Anti-counterfeit Hardware", "Secure disposal or reuse of equipment",
            "Access-based network segmentation", "Vulnerability Management", "Malware Protection",
            "ASIC/FPGA Manufacturing"
        ],
        "Physical/Logical Attack": [
            "A tamper resistant body", "Satellite Unit RF Encryption", "Traffic Flow Security",
            "Power Masking", "Secure disposal or reuse of equipment", "Access-based network segmentation",
            "Information classification and labelling", "Vulnerability Management", "Malware Protection"
        ],
        "Interception/Eavesdropping": [
            "Communications Security", "Satellite Unit RF Encryption", "Traffic Flow Security",
            "Power Masking", "Access-based network segmentation", "Information classification and labelling"
        ],
        "Jamming": [
            "Resilient Position Navigation and Timing", "Communication Physical Medium Space-Based",
            "Radio Frequency Mapping", "Antenna Nulling and Adaptive Filtering",
            "Defensive Jamming and Spoofing", "Emergency power sources",
            "Real-time physics model-based system verification"
        ],
        "Denial-of-Service": [
            "Security of Power Systems", "System redundancy", "Incident Recovery Plan",
            "Emergency power sources", "Traffic Flow Security",
            "Critical Services Delivery Requirements"
        ],
        "Masquerade/Spoofing": [
            "OSAM Dual Authorization", "Multi factor authentication", "Relay Protection",
            "Smart Contracts", "Resilient Position Navigation and Timing"
        ],
        "Replay": [
            "Relay Protection", "Satellite Unit RF Encryption", "On-board Message Encryption",
            "Session Termination", "Real-time physics model-based system verification"
        ],
        "Software Threats": [
            "Coding Standard", "Malware Protection", "Vulnerability scanning", "Vulnerability Management",
            "Software Updates", "Dynamic Code Analysis", "Static Code Analysis", "Process ID whitelisting",
            "Software Bill of Materials"
        ],
        "Unauthorized Access/Hijacking": [
            "Access rights", "Identity management", "Remote access management", "Multi factor authentication",
            "Access-based network segmentation", "Backdoor Commands"
        ],
        "Tainted hardware components": [
            "Anti-counterfeit Hardware", "ASIC/FPGA Manufacturing", "Tamper Protection",
            "Supplier Security Management"
        ],
        "Supply Chain": [
            "Supplier Security Management", "Software Bill of Materials", "Software Supply Chain Integrity",
            "Outsourced development", "Cloud Cybersecurity Measures"
        ]    }
    
    # Available mission types
    MISSION_TYPES = [
        "Insert type of mission",
        "Earth Observation Mission",
        "Communication Satellite", 
        "Scientific Mission",
        "Navigation Satellite",
        "On-Orbit Service"
    ]
    def __init__(self, root):
        self.root = root
        self.root.title("Risk Assessment Tool - Phase 0/A")
        self.root.state('zoomed')        
        self.root.configure(bg=self.COLORS['white'])
        # Data for threats and calculations
        self.threat_data = {}  # Saved data for threat
        self.combo_vars = {}   # ComboBox variables
        self.impact_entries = {}  # Table widgets
        
        # Variable for mission type
        self.mission_type_var = tk.StringVar(value=self.MISSION_TYPES[0])
        
        self.create_interface()
    def create_interface(self):
        """Creates the main interface"""
        # Header
        header = tk.Frame(self.root, bg=self.COLORS['light'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="Risk Assessment Tool - Phase 0/A", 
                font=('Segoe UI', 16, 'bold'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(pady=15)
        
        # Container principale
        main_frame = tk.Frame(self.root, bg=self.COLORS['white'])
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Threat table
        self.create_threats_table(main_frame)
        # Buttons
        self.create_buttons(main_frame)
    
    def create_threats_table(self, parent):
        """Creates the threat table"""
        # Main container
        main_container = tk.Frame(parent, bg=self.COLORS['white'])
        main_container.pack(fill='both', expand=True, pady=(0, 20))

        # Mission type selector (separate from the table)
        mission_frame = tk.LabelFrame(main_container, text="Mission Configuration",
                                     font=('Segoe UI', 11, 'bold'),
                                     bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                     padx=15, pady=10)
        mission_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(mission_frame, text="Mission Type:",
                font=('Segoe UI', 10, 'bold'),
                bg=self.COLORS['white'], fg=self.COLORS['dark']).pack(anchor='w')
        
        mission_combo = ttk.Combobox(mission_frame,
                                   textvariable=self.mission_type_var,
                                   values=self.MISSION_TYPES,
                                   font=('Segoe UI', 10),
                                   state='readonly')
        mission_combo.pack(fill='x', pady=(5, 0))

        # Threat table (separate from the mission selector)
        table_frame = tk.LabelFrame(main_container, text="Threat Risk Levels",
                                   font=('Segoe UI', 12, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=20, pady=15)
        table_frame.pack(fill='both', expand=True)
        
        # Header
        headers = ["Threat", "Risk Level"]
        for j, header in enumerate(headers):
            cell = tk.Label(table_frame, text=header,
                           font=('Segoe UI', 11, 'bold'),
                           bg=self.COLORS['primary'], fg=self.COLORS['white'],
                           relief='ridge', bd=1)
            cell.grid(row=0, column=j, sticky='ew', padx=1, pady=1, ipady=8)
        
        # Data Rows
        self.threat_cells = {}
        for i, threat in enumerate(self.THREATS, 1):
            # Threat name
            name_cell = tk.Label(table_frame, text=threat,
                               font=('Segoe UI', 10),
                               bg=self.COLORS['white'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, anchor='w')
            name_cell.grid(row=i, column=0, sticky='ew', padx=1, pady=1, ipady=5)
            
            # Risk level
            risk_cell = tk.Label(table_frame, text="",
                               font=('Segoe UI', 10),
                               bg=self.COLORS['white'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1)
            risk_cell.grid(row=i, column=1, sticky='ew', padx=1, pady=1, ipady=5)
            
            self.threat_cells[threat] = risk_cell

        # Grid configuration
        for j in range(2):
            table_frame.grid_columnconfigure(j, weight=1)
    
    def create_buttons(self, parent):
        """Creates the buttons"""
        button_frame = tk.Frame(parent, bg=self.COLORS['white'])
        button_frame.pack(fill='x')

        # ADD THREAT button
        add_btn = tk.Button(button_frame, text="ADD THREAT",
                           font=('Segoe UI', 12, 'bold'),
                           bg=self.COLORS['primary'], fg=self.COLORS['white'],
                           relief='flat', padx=30, pady=10,
                           command=self.open_threat_window)
        add_btn.pack(pady=10)

        # Buttons Word Export/Import
        bottom_frame = tk.Frame(button_frame, bg=self.COLORS['white'])
        bottom_frame.pack()
        
        if DOCX_AVAILABLE:
            export_word_btn = tk.Button(bottom_frame, text="EXPORT WORD",
                                       font=('Segoe UI', 11, 'bold'),
                                       bg='#2e86de', fg=self.COLORS['white'],
                                       relief='flat', padx=20, pady=8,
                                       command=self.export_to_word)
            export_word_btn.pack(side='left', padx=(0, 10))
            
            import_word_btn = tk.Button(bottom_frame, text="IMPORT WORD",
                                       font=('Segoe UI', 11, 'bold'),
                                       bg='#8e44ad', fg=self.COLORS['white'],
                                       relief='flat', padx=20, pady=8,
                                       command=self.import_from_word)
            import_word_btn.pack(side='left', padx=(10, 0))
        else:
            no_docx_label = tk.Label(bottom_frame, text="Word export/import unavailable - install python-docx",
                                   font=('Segoe UI', 9),
                                   bg=self.COLORS['white'], fg=self.COLORS['gray'])
            no_docx_label.pack()
        
        # Import Legacy Report button (always available)
        import_legacy_btn = tk.Button(bottom_frame, text="IMPORT LEGACY",
                                     font=('Segoe UI', 11, 'bold'),
                                     bg='#e67e22', fg=self.COLORS['white'],
                                     relief='flat', padx=20, pady=8,
                                     command=self.import_legacy_report)
        import_legacy_btn.pack(side='left', padx=(10, 0))
    
    def open_threat_window(self):
        """Open Threat Analysis window"""
        window = tk.Toplevel(self.root)
        window.title("Threat Analysis")
        window.geometry("1400x800")
        window.configure(bg=self.COLORS['white'])
        window.transient(self.root)
        window.grab_set()
        
        # Header
        header = tk.Frame(window, bg=self.COLORS['light'], height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="Threat Analysis - Asset Assessment",
                font=('Segoe UI', 14, 'bold'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(pady=12)

        # Main container with scroll
        self.create_threat_content(window)
    
    def create_threat_content(self, window):
        """Creates the threat content window"""
        # Scrollable canvas
        canvas = tk.Canvas(window, bg=self.COLORS['white'])
        scrollbar = tk.Scrollbar(window, orient="vertical", command=canvas.yview)
        content_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
        scrollbar.pack(side="right", fill="y")

        # Criteria table
        self.create_criteria_table(content_frame)
        
        # Threat selection
        threat_frame = tk.Frame(content_frame, bg=self.COLORS['white'])
        threat_frame.pack(fill='x', pady=(20, 20))
        
        tk.Label(threat_frame, text="Select Threat:",
                font=('Segoe UI', 11, 'bold'),
                bg=self.COLORS['white'], fg=self.COLORS['dark']).pack(anchor='w')
        
        self.selected_threat_var = tk.StringVar()
        threat_combo = ttk.Combobox(threat_frame,
                                   textvariable=self.selected_threat_var,
                                   values=self.THREATS,
                                   font=('Segoe UI', 10),
                                   state='readonly')
        threat_combo.pack(fill='x', pady=(5, 0))
        threat_combo.bind('<<ComboboxSelected>>', self.load_threat_data)
        
        # Asset table
        self.create_asset_table(content_frame)

        # Buttons frame
        buttons_frame = tk.Frame(content_frame, bg=self.COLORS['white'])
        buttons_frame.pack(pady=20)
        
        # Save button
        save_btn = tk.Button(buttons_frame, text="SAVE ASSESSMENT",
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['success'], fg=self.COLORS['white'],
                            relief='flat', padx=25, pady=10,
                            command=lambda: self.save_threat_assessment(window))
        save_btn.pack(side='left', padx=(0, 10))
        
        # Help button
        help_btn = tk.Button(buttons_frame, text="❓ Help",
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['gray'], fg=self.COLORS['white'],
                            relief='flat', padx=20, pady=10,
                            command=self.show_help)
        help_btn.pack(side='left')

        # Scroll with mouse wheel
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Bind mouse wheel to canvas and content frame
        canvas.bind("<MouseWheel>", on_mousewheel)
        content_frame.bind("<MouseWheel>", on_mousewheel)
        
        # Also bind to window itself to ensure it works when focused
        window.bind("<MouseWheel>", on_mousewheel)
            
    def create_criteria_table(self, parent):
        """Creates the assessment criteria table"""
        criteria_container = tk.LabelFrame(parent, 
                                         text="Assessment Criteria",
                                         font=('Segoe UI', 12, 'bold'),
                                         bg=self.COLORS['white'], 
                                         fg=self.COLORS['primary'], 
                                         padx=20, 
                                         pady=15,                                         
                                         relief='ridge', 
                                         bd=2)
        criteria_container.pack(fill='x', pady=(0, 20))

        # Creates the cells of the criteria table
        for i, row in enumerate(self.CRITERIA_DATA):
            for j, cell_text in enumerate(row):                    
                if i == 0:  # Header row                    
                    cell = tk.Label(criteria_container, text=cell_text,
                                   font=('Segoe UI', 10, 'bold'),
                                   bg=self.COLORS['criteria_header'], 
                                   fg=self.COLORS['white'],
                                   relief='ridge', 
                                   bd=1,                                   
                                   anchor='center',
                                   justify='center',
                                   wraplength=180,  
                                   width=22,        
                                   height=3,
                                   padx=3,         
                                   pady=2)
                else:  # Data rows
                    
                    font_weight = 'bold' if j == 0 else 'normal'                        
                    cell = tk.Label(criteria_container, text=cell_text,
                                   font=('Segoe UI', 9, font_weight),
                                   bg=self.COLORS['criteria_bg'],
                                   fg=self.COLORS['dark'],
                                   relief='ridge',
                                   bd=1,
                                   anchor='nw',
                                   justify='left',
                                   wraplength=180,  
                                   width=22,        
                                   height=4,        
                                   padx=6,          
                                   pady=3)          
                
                cell.grid(row=i, column=j, padx=2, pady=2, sticky='ew', ipady=5)

        # Grid configuration with uniform column sizes
        for j in range(6):
            criteria_container.grid_columnconfigure(j, weight=1, minsize=160, uniform="criteria_cols")  
        
        for i in range(6):  # 5 data rows + 1 header
            criteria_container.grid_rowconfigure(i, minsize=60, uniform="criteria_rows")  
    
    def create_asset_table(self, parent):
        """Creates the asset assessment table"""
        table_frame = tk.LabelFrame(parent, text="Asset Assessment (Values 1-5)",
                                   font=('Segoe UI', 11, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=15, pady=15)        
        table_frame.pack(fill='both', expand=True)
        
        # Headers
        headers = ["Category", "Sub-Category", "Vulnerability", "Access", "Defense", 
                  "Operational Impact", "Recovery", "Likelihood", "Impact", "Risk"]
        
        for j, header in enumerate(headers):
            cell = tk.Label(table_frame, text=header,
                           font=('Segoe UI', 10, 'bold'),
                           bg=self.COLORS['primary'], fg=self.COLORS['white'],
                           relief='ridge', bd=1, width=12)
            cell.grid(row=0, column=j, padx=1, pady=1, sticky='ew')
        
        # Reset storage
        self.impact_entries = {}
        self.combo_vars = {}
        
        # Assessment Rows (9 Rows: 1 for each category)
        for i in range(9):
            category, subcategory = self.ASSET_CATEGORIES[i]
            asset_key = f"{i+1}_probability"  # Unique key for asset
            # Category (read-only)
            cat_cell = tk.Label(table_frame, text=category,
                               font=('Segoe UI', 9, 'bold'),
                               bg=self.COLORS['light'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, width=10)
            cat_cell.grid(row=i+1, column=0, padx=1, pady=1, sticky='ew')
            
            # Sub-Category (read-only)
            sub_cell = tk.Label(table_frame, text=subcategory,
                               font=('Segoe UI', 9),
                               bg=self.COLORS['light'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, width=15)
            sub_cell.grid(row=i+1, column=1, padx=1, pady=1, sticky='ew')

            # Storage for this row
            row_entries = {}
            self.combo_vars[asset_key] = {}
            
            # Writable columns (2-6: Vulnerability, Access, Defense, Operational Impact, Recovery)
            for j in range(2, 7):
                combo_var = tk.StringVar(value="")                
                combo = ttk.Combobox(table_frame,
                                    textvariable=combo_var,
                                    values=["", "1", "2", "3", "4", "5"],
                                    font=('Segoe UI', 9),
                                    width=8, state='readonly')
                combo.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                
                row_entries[j-2] = combo  # 0-based index
                self.combo_vars[asset_key][j-2] = combo_var

                # Bind calculations
                if j <= 4:  # Vulnerability, Access, Defense -> Likelihood
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.calculate_likelihood(key))
                elif j <= 6:  # Operational Impact, Recovery -> Impact
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.calculate_impact(key))
            # Colonne calcolate (7-9: Likelihood, Impact, Risk) - read-only
            for j in range(7, 10):
                calc_cell = tk.Label(table_frame, text="",
                                   font=('Segoe UI', 9),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   relief='ridge', bd=1, width=10)
                calc_cell.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                row_entries[j-2] = calc_cell
            
            self.impact_entries[asset_key] = row_entries        
        # Column 0 (Category):
        table_frame.grid_columnconfigure(0, weight=1, minsize=120, uniform="small_cols")
        # Column 1 (Sub-Category): 
        table_frame.grid_columnconfigure(1, weight=1, minsize=180, uniform="sub_category_col")
        # Columns 2-9:
        for j in range(2, 10):
            table_frame.grid_columnconfigure(j, weight=1, minsize=120, uniform="small_cols")
        
        for i in range(10):  # 9 data rows + 1 header
            table_frame.grid_rowconfigure(i, minsize=40, uniform="rows")
    
    def calculate_likelihood(self, key):
        """Calculates Likelihood using quadratic mean of three criteria"""
        if key not in self.combo_vars:
            return

        # Get values Vulnerability, Access, Defense (columns 0,1,2)
        values = []
        for col_idx in [0, 1, 2]:
            if col_idx not in self.combo_vars[key]:
                self.update_display(key, 5, "")  # Likelihood column
                return
            
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                self.update_display(key, 5, "")
                return
            
            try:
                values.append(float(value_str))
            except ValueError:
                self.update_display(key, 5, "")
                return
        
        if len(values) != 3:
            self.update_display(key, 5, "")
            return

        # Calculate Likelihood using quadratic mean
        quadratic_mean = math.sqrt(sum(x**2 for x in values) / 3)
        likelihood = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
        likelihood = max(0.0, min(1.0, likelihood))

        # Update display with category instead of numeric value
        likelihood_category = self.value_to_category(likelihood)
        self.update_display(key, 5, likelihood_category)  # Likelihood column

        # Recalculate risk if Impact is also available
        self.calculate_risk(key)

        # Update main table in real-time also for likelihood
        self.update_main_table_risk_realtime()
    def get_saved_likelihood(self, threat_name, asset_num):
        """Get saved likelihood for specific threat/asset using only base calculation"""
        if threat_name not in self.threat_data:
            return 0.0
        
        asset_key = f"{asset_num}_probability"
        if asset_key not in self.threat_data[threat_name]:
            return 0.0
        
        row_data = self.threat_data[threat_name][asset_key]
        
        try:
            # Check values for Vulnerability, Access, Defense
            if not isinstance(row_data, dict):
                return 0.0
            
            values = []
            for i in [0, 1, 2]:  # Vulnerability, Access, Defense
                if str(i) not in row_data:
                    return 0.0
                val = row_data[str(i)]
                if not val or val == "0":
                    return 0.0
                values.append(float(val))
            
            if len(values) == 3:
                # Calculate base likelihood using quadratic mean
                quadratic_mean = math.sqrt(sum(x**2 for x in values) / 3)
                likelihood = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
                return max(0.0, min(1.0, likelihood))
        
        except (ValueError, KeyError, TypeError):
            pass
        
        return 0.0
    def calculate_impact(self, key):
        """Calculates Impact as the quadratic mean of Operational Impact and Recovery Time"""
        if key not in self.combo_vars:
            return

        # Get Operational Impact, Recovery values (columns 3,4)
        values = []
        for col_idx in [3, 4]:
            if col_idx not in self.combo_vars[key]:
                self.update_display(key, 6, "")  # Impact column
                return
            
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                self.update_display(key, 6, "")
                return
            
            try:
                values.append(float(value_str))
            except ValueError:
                self.update_display(key, 6, "")
                return
        if len(values) == 2:
            # Quadratic mean normalized
            quadratic_mean = math.sqrt(sum(x**2 for x in values) / 2)
            impact = (quadratic_mean - 1) / 4  # [1,5] -> [0,1]
            impact = max(0.0, min(1.0, impact))

            # Update display with category instead of numeric value
            impact_category = self.value_to_category(impact)
            self.update_display(key, 6, impact_category)  # Impact column

            # Recalculate risk
            self.calculate_risk(key)
        else:
            self.update_display(key, 6, "")
    def calculate_risk(self, key):
        """Calculates Risk using the Likelihood x Impact matrix"""
        if key not in self.impact_entries:
            return

        # Get Likelihood and Impact (now they are already categories)
        likelihood_widget = self.impact_entries[key][5]  # Likelihood column
        impact_widget = self.impact_entries[key][6]      # Impact column
        
        likelihood_cat = likelihood_widget.cget('text') if hasattr(likelihood_widget, 'cget') else ""
        impact_cat = impact_widget.cget('text') if hasattr(impact_widget, 'cget') else ""
        
        if not likelihood_cat or not impact_cat:
            self.update_display(key, 7, "")  # Risk column
            return

        # Check if they are valid categories
        valid_categories = ["Very Low", "Low", "Medium", "High", "Very High"]
        if likelihood_cat in valid_categories and impact_cat in valid_categories:
            # Get risk from matrix
            risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
            self.update_display(key, 7, risk_level)  # Risk column

            # Update main table in real-time
            self.update_main_table_risk_realtime()
        else:
            self.update_display(key, 7, "")
    
    def value_to_category(self, value):
        """Converts numeric value to category"""
        if value <= 0.1:
            return "Very Low"
        elif value <= 0.4:
            return "Low"
        elif value <= 0.7:
            return "Medium"
        elif value <= 0.9:
            return "High"
        else:
            return "Very High"
    
    def update_display(self, key, col_index, value):
        """Updates the display of a cell"""
        if key in self.impact_entries and col_index in self.impact_entries[key]:
            widget = self.impact_entries[key][col_index]
            if hasattr(widget, 'config'):
                widget.config(text=value)
    
    def load_threat_data(self, event=None):
        """Loads data for selected threat"""
        selected_threat = self.selected_threat_var.get()

        # Clear all fields
        for key in self.impact_entries:
            self.clear_asset_data(key)

        # Load data if it exists
        if selected_threat and selected_threat in self.threat_data:
            threat_data = self.threat_data[selected_threat]
            
            for key, row_data in threat_data.items():
                if key in self.combo_vars:
                    for col_idx, value in row_data.items():
                        if int(col_idx) in self.combo_vars[key]:
                            self.combo_vars[key][int(col_idx)].set(value)

        # Recalculate everything
        for key in self.impact_entries:
            self.calculate_likelihood(key)
            self.calculate_impact(key)
    
    def clear_asset_data(self, key):
        """Clears data for an asset"""
        if key in self.combo_vars:
            for combo_var in self.combo_vars[key].values():
                combo_var.set("")
        
        if key in self.impact_entries:
            for col_idx in [5, 6, 7]:  # Likelihood, Impact, Risk
                if col_idx in self.impact_entries[key]:
                    self.update_display(key, col_idx, "")
    def save_threat_assessment(self, window):
        """Saves current threat assessment"""
        selected_threat = self.selected_threat_var.get()
        if not selected_threat:
            messagebox.showwarning("Warning", "Please select a threat first!")
            return

        # Collect data
        threat_data = {}
        for key in self.combo_vars:
            row_data = {}
            for col_idx, combo_var in self.combo_vars[key].items():
                value = combo_var.get().strip()
                if value:
                    row_data[str(col_idx)] = value
            
            if row_data:
                threat_data[key] = row_data
        
        # Save data
        self.threat_data[selected_threat] = threat_data

        # Update main table with maximum risks of ALL threats
        self.update_all_threats_in_main_table()
        
        messagebox.showinfo("Success", f"Assessment for '{selected_threat}' saved!")
        window.destroy()
    
    def update_main_table_risk(self, threat_name):
        """Updates main table with maximum risk"""
        if threat_name not in self.threat_data or threat_name not in self.threat_cells:
            return

        # Find maximum risk among all assets
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}
        max_risk = ""
        max_priority = 0

        # Check all assets
        for key in self.impact_entries:
            risk_widget = self.impact_entries[key].get(7)  # Risk column
            if risk_widget and hasattr(risk_widget, 'cget'):
                risk_text = risk_widget.cget('text')
                priority = risk_priorities.get(risk_text, 0)
                if priority > max_priority:
                    max_priority = priority
                    max_risk = risk_text        # Update main table
        self.threat_cells[threat_name].config(text=max_risk)
    
    def update_all_threats_in_main_table(self):
        """Updates main table with maximum risks of all saved threats"""
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}

        # For each threat that has saved data
        for threat_name in self.threat_data.keys():
            if threat_name not in self.threat_cells:
                continue
            
            max_risk = ""
            max_priority = 0

            # Calculate maximum risk for this threat
            threat_data = self.threat_data[threat_name]
            
            for asset_key, asset_data in threat_data.items():
                # Calculate likelihood for this asset
                likelihood = self.calculate_likelihood_from_saved_data(threat_name, asset_key, asset_data)

                # Calculate impact for this asset
                impact = self.calculate_impact_from_saved_data(asset_data)

                # Calculate risk if both are available
                if likelihood >= 0 and impact >= 0:
                    likelihood_cat = self.value_to_category(likelihood)
                    impact_cat = self.value_to_category(impact)
                    risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
                    
                    priority = risk_priorities.get(risk_level, 0)
                    if priority > max_priority:
                        max_priority = priority
                        max_risk = risk_level

            # Update main table for this threat
            self.threat_cells[threat_name].config(text=max_risk)

    def calculate_likelihood_from_saved_data(self, threat_name, asset_key, asset_data):
        """Calculates likelihood from saved data using quadratic mean"""
        try:
            # Check if necessary values are present (Vulnerability, Access, Defense)
            if not all(str(i) in asset_data for i in [0, 1, 2]):
                return 0.0
            
            values = []
            for i in [0, 1, 2]:
                val = asset_data[str(i)]
                if not val or val == "0":
                    return 0.0
                values.append(float(val))
            
            if len(values) == 3:
                # Calculate likelihood using quadratic mean
                quadratic_mean = math.sqrt(sum(x**2 for x in values) / 3)
                likelihood = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
                return max(0.0, min(1.0, likelihood))
        
        except (ValueError, KeyError, TypeError):
            pass
        
        return 0.0
    def calculate_impact_from_saved_data(self, asset_data):
        """Calculates impact from saved data using quadratic mean"""
        try:
            # Check if necessary values are present (Operational Impact, Recovery)
            if not all(str(i) in asset_data for i in [3, 4]):
                return 0.0
            
            values = []
            for i in [3, 4]:  # Operational Impact, Recovery
                val = asset_data[str(i)]
                if not val or val == "0":
                    return 0.0
                values.append(float(val))
            
            if len(values) == 2:
                # Quadratic mean normalized
                quadratic_mean = math.sqrt(sum(x**2 for x in values) / 2)
                impact = (quadratic_mean - 1) / 4  # [1,5] -> [0,1]
                return max(0.0, min(1.0, impact))
        
        except (ValueError, KeyError, TypeError):
            pass
        
        return 0.0
    
    def update_main_table_risk_realtime(self):
        """Updates main table in real-time during calculations"""
        current_threat = self.selected_threat_var.get()
        if not current_threat or current_threat not in self.threat_cells:
            return
        
        # Find maximum risk among all currently displayed assets
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}
        max_risk = ""
        max_priority = 0
        
        # Check all currently displayed assets
        for key in self.impact_entries:
            risk_widget = self.impact_entries[key].get(7)  # Risk column
            if risk_widget and hasattr(risk_widget, 'cget'):
                risk_text = risk_widget.cget('text')
                priority = risk_priorities.get(risk_text, 0)
                if priority > max_priority:
                    max_priority = priority
                    max_risk = risk_text

        # Update main table in real time
        if max_risk:
            self.threat_cells[current_threat].config(text=max_risk)
    
    def get_max_risk_for_threat(self, threat_name):
        """Calculates the maximum risk for a specific threat"""
        if threat_name not in self.threat_data:
            return ""
        
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}
        max_risk = ""
        max_priority = 0

        # Calculate maximum risk for this threat
        threat_data = self.threat_data[threat_name]
        
        for asset_key, asset_data in threat_data.items():
            # Calculate likelihood for this asset
            likelihood = self.calculate_likelihood_from_saved_data(threat_name, asset_key, asset_data)

            # Calculate impact for this asset
            impact = self.calculate_impact_from_saved_data(asset_data)

            # Calculate risk if both are available
            if likelihood >= 0 and impact >= 0:
                likelihood_cat = self.value_to_category(likelihood)
                impact_cat = self.value_to_category(impact)
                risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
                
                priority = risk_priorities.get(risk_level, 0)
                if priority > max_priority:
                    max_priority = priority
                    max_risk = risk_level
        
        return max_risk

    # ===== WORD EXPORT/IMPORT METHODS =====
    def export_to_word(self):
        """Exports the risk assessment to a Word document"""
        if not DOCX_AVAILABLE:
            messagebox.showerror("Error", "python-docx library not available!\nInstall with: pip install python-docx")
            return
            
        try:
            # Mission type
            mission_type = self.mission_type_var.get()
            if mission_type == self.MISSION_TYPES[0]:  
                mission_type = ""  # Do not show if it is the default value

            # Automatic file name with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"Risk_Assessment_0-A_{timestamp}.docx"

            # Create Output directory if it doesn't exist
            output_dir = os.path.join(os.path.dirname(__file__), "Output")
            os.makedirs(output_dir, exist_ok=True)

            # Destination folder (Output directory)
            file_path = os.path.join(output_dir, filename)

            # Create Word document
            doc = Document()
            
            # Add content
            self.add_word_title(doc)
            if mission_type:
                self.add_mission_type(doc, mission_type)
            self.add_main_threats_table(doc)
            self.add_threat_details(doc)

            # Save document
            doc.save(file_path)
            
            messagebox.showinfo("Success", f"Risk Assessment exported to:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error exporting to Word:\n{str(e)}")
    def import_from_word(self):
        """Import data from a previously exported Word document"""
        if not DOCX_AVAILABLE:
            messagebox.showerror("Error", "python-docx library not available!\nInstall with: pip install python-docx")
            return
            
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Word documents", "*.docx")],
                title="Import Risk Assessment from Word"
            )
            
            if not file_path:
                return            # Load Word document
            doc = Document(file_path)

            # Extract data
            self.extract_mission_type_from_word(doc)
            self.extract_threats_data_from_word(doc)

            # Update interface
            self.update_all_threats_in_main_table()
            
            messagebox.showinfo("Success", f"Risk Assessment imported from:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error importing from Word:\n{str(e)}")
    
    def import_legacy_report(self):
        """Import data from a legacy Word report"""
        if not DOCX_AVAILABLE:
            messagebox.showerror("Error", "python-docx library not available!\nInstall with: pip install python-docx")
            return
            
        try:
            file_path = filedialog.askopenfilename(
                filetypes=[("Word documents", "*.docx"), ("All files", "*.*")],
                title="Import Legacy Risk Assessment Report"
            )
            
            if not file_path:
                return
            
            # Load Word document
            doc = Document(file_path)
            
            # Parse the legacy report
            threats_data = self.parse_legacy_word_report(doc)
            
            if not threats_data:
                messagebox.showwarning("Warning", "No threat data found in the legacy report.")
                return
            
            # Clear existing data
            self.threat_data = {}
            
            # Import the parsed data into our data structure
            for threat_name, threat_info in threats_data.items():
                if threat_name in self.THREATS:
                    likelihood = threat_info['likelihood']
                    asset_categories = threat_info['assets']
                    
                    # Create threat data structure
                    threat_data = {}
                    
                    # For each asset category mentioned in the legacy report
                    for asset_category in asset_categories:
                        # Find the corresponding asset index in our ASSET_CATEGORIES
                        for i, (main_cat, sub_cat) in enumerate(self.ASSET_CATEGORIES):
                            if asset_category in sub_cat or asset_category in main_cat:
                                asset_key = f"{i+1}_probability"
                                
                                # Set minimal data to create a valid likelihood
                                # We'll use default values for Vulnerability, Access, Defense
                                # but set them to produce the desired likelihood
                                likelihood_val = self.category_to_value(likelihood)
                                base_val = self.likelihood_to_base_value(likelihood_val)
                                
                                # Use base_val - 1 for better calibration
                                adjusted_val = max(1, base_val)
                                
                                threat_data[asset_key] = {
                                    '0': str(adjusted_val),  # Vulnerability
                                    '1': str(adjusted_val),  # Access Control
                                    '2': str(adjusted_val),  # Defense Capability
                                    '3': '3',  # Operational Impact (medium)
                                    '4': '3'   # Recovery Time (medium)
                                }
                    
                    if threat_data:
                        self.threat_data[threat_name] = threat_data
            
            # Update interface
            self.update_all_threats_in_main_table()
            
            imported_count = len(threats_data)
            messagebox.showinfo("Success", f"Legacy report imported successfully!\n"
                                         f"Imported {imported_count} threats from:\n{file_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Error importing legacy report:\n{str(e)}")
    
    def parse_legacy_report(self, content):
        """Parse a legacy text report and extract threat probability data"""
        threats_data = {}
        
        try:
            # Look for the "Main Threats Overview" section
            lines = content.split('\n')
            in_main_threats_section = False
            
            for line in lines:
                line = line.strip()
                
                # Check if we're entering the Main Threats Overview section
                if 'Main Threats Overview' in line:
                    in_main_threats_section = True
                    continue
                
                # Check if we're leaving the main threats section (entering detailed analysis)
                if in_main_threats_section and ('Detailed Threat Analysis' in line or 'Risk Matrix' in line):
                    break
                
                # If we're in the main threats section, look for threat data
                if in_main_threats_section and '|' in line:
                    # Skip header lines
                    if 'Threat' in line and 'Probability' in line:
                        continue
                    if line.startswith('|--') or line.startswith('---'):
                        continue
                    
                    # Parse threat line: "| Threat Name | Probability |"
                    parts = [part.strip() for part in line.split('|') if part.strip()]
                    
                    if len(parts) >= 2:
                        threat_name = parts[0].strip()
                        probability = parts[1].strip()
                        
                        # Map probability values to our standard values
                        probability_mapping = {
                            'Very Low': 'Very Low',
                            'Low': 'Low',
                            'Medium': 'Medium',
                            'High': 'High',
                            'Very High': 'Very High',
                            'VL': 'Very Low',
                            'L': 'Low',
                            'M': 'Medium',
                            'H': 'High',
                            'VH': 'Very High'
                        }
                        
                        # Normalize probability
                        probability = probability_mapping.get(probability, probability)
                        
                        # Only add if it's a valid threat name from our list
                        if threat_name in self.THREATS and probability in ['Very Low', 'Low', 'Medium', 'High', 'Very High']:
                            threats_data[threat_name] = probability
            
        except Exception as e:
            print(f"Error parsing legacy report: {e}")
        
        return threats_data
    
    def add_word_title(self, doc):
        """Adds the title to the Word document"""
        # Main title
        title = doc.add_heading('Risk Assessment', 0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Data
        date_para = doc.add_paragraph(f'Generated on: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}')
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_paragraph()  # Empty space

    def add_mission_type(self, doc, mission_type=None):
        """Adds the mission type to the document"""
        if mission_type:
            mission_para = doc.add_paragraph(f'Mission Type: {mission_type}')
            mission_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            doc.add_paragraph()  # Empty space

    def add_main_threats_table(self, doc):
        """Adds the main threats table"""
        doc.add_heading('Main Threats Overview', level=1)

        # Create simple table: Threat | Risk Level (as in main window)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        # Header
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Threat'
        header_cells[1].text = 'Risk Level'

        # Format header
        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True
        # Add data for each threat (use the same calculations as the main table)
        for threat in self.THREATS:
            row_cells = table.add_row().cells
            row_cells[0].text = threat

            # Calculate maximum risk for this threat
            max_risk = self.get_max_risk_for_threat(threat)
            row_cells[1].text = max_risk if max_risk else ""
        
        doc.add_page_break()
    def add_threat_details(self, doc):
        """Adds details for each threat with risk"""
        doc.add_heading('Detailed Threat Analysis', level=1)
        
        threats_with_data = []

        # First find all threats that have data
        for threat in self.THREATS:
            if threat in self.threat_data and self.threat_data[threat]:
                max_risk = self.get_max_risk_for_threat(threat)
                if max_risk:  # Only if a risk has been calculated
                    threats_with_data.append(threat)
        
        if not threats_with_data:
            doc.add_paragraph("No threats with assessed risk data.")
            return        # Add details for each threat with data
        for threat in threats_with_data:
            # Threat title
            doc.add_heading(f'{threat}', level=2)

            # Add asset table for this threat
            self.add_threat_asset_table(doc, threat)

            # Add countermeasures table for this threat
            self.add_threat_countermeasures_table(doc, threat)
            doc.add_paragraph()  # Empty space between threats
        
        # Add reference tables at the end
        doc.add_page_break()
        self.add_criteria_reference_table(doc)
        self.add_risk_matrix_table(doc)

    def add_threat_asset_table(self, doc, threat_name):
        """Creates asset assessment table for a specific threat"""
        doc.add_heading(f'Asset Assessment for {threat_name}', level=3)
        
        table = doc.add_table(rows=1, cols=9)
        table.style = 'Table Grid'

        # Exact header for import
        header_cells = table.rows[0].cells
        headers = ['Asset', 'Vulnerability', 'Access Control', 'Defense Capability', 
                  'Operational Impact', 'Recovery Time', 'Likelihood', 'Impact', 'Risk Level']
        
        for i, header in enumerate(headers):
            header_cells[i].text = header
            header_cells[i].paragraphs[0].runs[0].bold = True

        # Add rows for each asset with data
        assets_added = 0
        if threat_name in self.threat_data:
            threat_data = self.threat_data[threat_name]
            
            for asset_key, asset_data in threat_data.items():
                # Extract asset index from key (e.g., "1_probability" -> 0)
                try:
                    asset_index = int(asset_key.split('_')[0]) - 1
                    if 0 <= asset_index < len(self.ASSET_CATEGORIES):
                        category, asset_name = self.ASSET_CATEGORIES[asset_index]

                        # Calculate likelihood and impact
                        likelihood = self.calculate_likelihood_from_saved_data(threat_name, asset_key, asset_data)
                        impact = self.calculate_impact_from_saved_data(asset_data)

                        # Only if we have valid data
                        if asset_data and (likelihood >= 0 or impact >= 0):
                            row_cells = table.add_row().cells

                            # Asset name (important: must exactly match ASSET_CATEGORIES)
                            row_cells[0].text = asset_name

                            # Criteria (columns 0-4 correspond to Vulnerability, Access, Defense, Operational Impact, Recovery)
                            criteria_keys = ['0', '1', '2', '3', '4']
                            for j, key in enumerate(criteria_keys):
                                if key in asset_data:
                                    score = asset_data[key]
                                    # Specific format for import
                                    row_cells[j + 1].text = f"Score {score}"
                                else:
                                    row_cells[j + 1].text = "N/A"

                            # Calculate likelihood and impact if possible
                            if likelihood >= 0:
                                likelihood_cat = self.value_to_category(likelihood)
                                row_cells[6].text = likelihood_cat
                            else:
                                row_cells[6].text = "N/A"
                            
                            if impact >= 0:
                                impact_cat = self.value_to_category(impact)
                                row_cells[7].text = impact_cat
                            else:
                                row_cells[7].text = "N/A"
                            
                            # Risk Level
                            if likelihood >= 0 and impact >= 0:
                                likelihood_cat = self.value_to_category(likelihood)
                                impact_cat = self.value_to_category(impact)
                                risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "N/A")
                                row_cells[8].text = risk_level
                            else:
                                row_cells[8].text = "N/A"
                            
                            assets_added += 1
                            
                except (ValueError, IndexError) as e:
                    print(f"Error processing asset {asset_key}: {e}")
                    continue
        
        # If no assets have data, add a placeholder row
        if assets_added == 0:
            row_cells = table.add_row().cells
            row_cells[0].text = "No asset data available"
            for i in range(1, 9):
                row_cells[i].text = "N/A"
        
        print(f"Added {assets_added} assets to table for threat {threat_name}")
        doc.add_paragraph()  # Space after the table
    
    def add_threat_countermeasures_table(self, doc, threat_name):
        """Add a table with security check for a specific threat"""
        # Add subtitle for controls
        doc.add_heading(f'Security Controls for {threat_name}', level=3)

        # Check if there are controls for this threat
        if threat_name not in self.THREAT_COUNTERMEASURES:
            doc.add_paragraph("No specific security controls defined for this threat.")
            return
        
        countermeasures = self.THREAT_COUNTERMEASURES[threat_name]

        # Create table with 2 columns (Control #, Control Name)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        
        # Header
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Control #'
        header_cells[1].text = 'Security Control'

        # Format header
        for cell in header_cells:
            cell.paragraphs[0].runs[0].bold = True

        # Add controls
        for i, control in enumerate(countermeasures, 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(i)
            row_cells[1].text = control

        # Add space after the table
        doc.add_paragraph()
    def extract_mission_type_from_word(self, doc):
        """Extracts the mission type from the Word document and updates the dropdown"""
        for paragraph in doc.paragraphs:
            text = paragraph.text.strip()
            if text.startswith("Mission Type:"):
                mission_type = text.replace("Mission Type:", "").strip()
                # Update the dropdown if the mission type is in the list
                if mission_type in self.MISSION_TYPES:
                    self.mission_type_var.set(mission_type)
                elif mission_type:  # If not in the list but not empty, use the first as fallback
                    self.mission_type_var.set(mission_type)
                    break
    def extract_threats_data_from_word(self, doc):
        """Extracts threat data from the Word document from the Detailed Threat Analysis section"""
        in_detailed_section = False

        # First, reset existing data to avoid conflicts
        self.threat_data = {}

        # Improved method: scan all document elements in order
        all_elements = []
        for element in doc.element.body:
            if element.tag.endswith('p'):  # Paragraph
                para_text = ""
                for para in doc.paragraphs:
                    if para._element == element:
                        para_text = para.text.strip()
                        break
                all_elements.append(('paragraph', para_text))
            elif element.tag.endswith('tbl'):  # Table
                for table in doc.tables:
                    if table._element == element:
                        all_elements.append(('table', table))
                        break

        # Now process elements in order
        current_threat = None
        threat_table_count = {}  # Count tables for each threat

        for element_type, element_data in all_elements:
            if element_type == 'paragraph':
                text = element_data

                # Check if we are in the "Detailed Threat Analysis" section
                if "Detailed Threat Analysis" in text:
                    in_detailed_section = True
                    print("✅ Found Detailed Threat Analysis section")
                    continue

                # If we are in the detailed section, look for threat names
                if in_detailed_section and text in self.THREATS:
                    current_threat = text
                    threat_table_count[current_threat] = 0
                    print(f"📋 Found threat: {current_threat}")
                    
            elif element_type == 'table' and current_threat and in_detailed_section:
                table = element_data
                threat_table_count[current_threat] += 1
                table_number = threat_table_count[current_threat]
                
                print(f"🔍 Processing table #{table_number} for threat: {current_threat}")

                # Check table type by number of columns
                if len(table.columns) == 9:
                    # Asset table
                    print(f"   → Asset table detected (9 columns)")
                    self.extract_asset_table_data(table, current_threat)
                elif len(table.columns) == 2:
                    # Controls table (ignore for data import)
                    print(f"   → Controls table detected (2 columns) - skipping")
                else:
                    print(f"   → Unknown table format ({len(table.columns)} columns) - skipping")
                    
        print(f"🎯 Import completed. Found data for threats: {list(self.threat_data.keys())}")

        # Debug: show imported data
        for threat_name, threat_data in self.threat_data.items():
            print(f"   {threat_name}: {len(threat_data)} assets")
      
    def extract_asset_table_data(self, table, threat_name):
        """Extracts data from the asset table for a specific threat"""
        try:
            print(f"🔍 Extracting asset table data for threat: {threat_name}")
            print(f"   Table dimensions: {len(table.rows)} rows × {len(table.columns)} columns")

            # Check table format (must have 9 columns)
            if len(table.columns) != 9:
                print(f"❌ Invalid table format: expected 9 columns, got {len(table.columns)}")
                return

            # Check header to confirm it's the right table
            header_row = table.rows[0]
            expected_headers = ['Asset', 'Vulnerability', 'Access Control', 'Defense Capability', 
                              'Operational Impact', 'Recovery Time', 'Likelihood', 'Impact', 'Risk Level']
            
            header_match = True
            for i, expected in enumerate(expected_headers):
                if i < len(header_row.cells):
                    cell_text = header_row.cells[i].text.strip()
                    if expected.lower() not in cell_text.lower():
                        header_match = False
                        break
            
            if not header_match:
                print(f"❌ Header mismatch - not an asset table")
                return
            
            print(f"✅ Valid asset table confirmed")

            # Initialized data for this threat if not exists
            if threat_name not in self.threat_data:
                self.threat_data[threat_name] = {}

            # Process each data row (skip header)
            data_rows_processed = 0
            for row_idx, row in enumerate(table.rows[1:], 1):
                try:
                    cells = row.cells
                    if len(cells) < 9:
                        print(f"⚠️  Row {row_idx}: insufficient cells ({len(cells)})")
                        continue

                    # Extract asset name from the first cell
                    asset_name = cells[0].text.strip()
                    if not asset_name:
                        print(f"⚠️  Row {row_idx}: empty asset name")
                        continue
                    
                    print(f"   Processing asset: '{asset_name}'")

                    # Find the index of the corresponding asset in the standard categories
                    asset_index = None
                    for i, (category, name) in enumerate(self.ASSET_CATEGORIES):
                        if name.lower() == asset_name.lower():
                            asset_index = i + 1  # Index start from 1
                            break
                    
                    if asset_index is None:
                        print(f"❌ Asset '{asset_name}' not found in standard categories")
                        # Try partial matching
                        for i, (category, name) in enumerate(self.ASSET_CATEGORIES):
                            if asset_name.lower() in name.lower() or name.lower() in asset_name.lower():
                                asset_index = i + 1
                                print(f"✅ Found partial match: '{name}' → using index {asset_index}")
                                break
                        
                        if asset_index is None:
                            continue

                    # Create key for threat data
                    asset_key = f"{asset_index}_probability"

                    # Extract scores from criteria (columns 1-5)
                    criteria_scores = {}
                    valid_scores = 0

                    for j in range(1, 6):  # Columns 1-5 for the 5 criteria
                        cell_text = cells[j].text.strip()

                        # Various parsing formats
                        score = self.parse_score_from_cell(cell_text)
                        
                        if score is not None:
                            criteria_scores[str(j-1)] = str(score)  # Save as string with index 0-4
                            valid_scores += 1
                            print(f"     Criterion {j-1}: {score}")
                        else:
                            print(f"     Criterion {j-1}: could not parse '{cell_text}'")

                    # Save only if we have at least 3 valid criteria (to calculate likelihood/impact)
                    if valid_scores >= 3:
                        self.threat_data[threat_name][asset_key] = criteria_scores
                        data_rows_processed += 1
                        print(f"✅ Saved data for asset {asset_index} ({asset_name}): {valid_scores} criteria")
                    else:
                        print(f"❌ Insufficient valid criteria ({valid_scores}/5) for asset {asset_name}")
                    
                except Exception as e:
                    print(f"❌ Error processing row {row_idx}: {e}")
                    continue
            
            print(f"🎯 Processed {data_rows_processed} valid asset rows for threat '{threat_name}'")
                    
        except Exception as e:
            print(f"❌ Error extracting asset table data for {threat_name}: {e}")
    
    def parse_score_from_cell(self, cell_text):
        """Extracts a score from a Word table cell with various formats"""
        if not cell_text:
            return None
        
        text = cell_text.strip()

        # Format 1: "Score X"
        if "score" in text.lower():
            try:
                # Remove "Score" and take the number
                score_str = text.lower().replace("score", "").strip()
                return int(score_str)
            except ValueError:
                pass

        # Format 2: Only number
        if text.isdigit():
            score = int(text)
            if 1 <= score <= 5:  # Valida range
                return score

        # Format 3: Number in a longer string
        import re
        numbers = re.findall(r'\b([1-5])\b', text)
        if numbers:
            return int(numbers[0])

        # Format 4: "N/A" or empty
        if text.lower() in ['n/a', 'na', '-', '']:
            return None
        
        return None

    def add_criteria_reference_table(self, doc):
        """Adds the assessment criteria reference table"""
        doc.add_heading('Assessment Criteria Reference', level=1)

        # Creates criteria table (6 columns: Criteria + 5 score levels)
        table = doc.add_table(rows=len(self.CRITERIA_DATA), cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Populates the table with criteria data
        for i, row_data in enumerate(self.CRITERIA_DATA):
            row_cells = table.rows[i].cells
            for j, cell_text in enumerate(row_data):
                row_cells[j].text = cell_text

                # Formats the header (first row)
                if i == 0:
                    row_cells[j].paragraphs[0].runs[0].bold = True                # Formats the first column (criteria names)
                elif j == 0:
                    row_cells[j].paragraphs[0].runs[0].bold = True

        doc.add_paragraph()  # Space after the table

    def add_risk_matrix_table(self, doc):
        """Adds the risk matrix ISO 27005"""
        doc.add_heading('Risk Assessment Matrix (ISO 27005)', level=1)

        # Defines the levels for the matrix (ISO 27005)
        levels = ["Very High", "High", "Medium", "Low", "Very Low"]

        # Creates 6x6 table (header + 5x5 matrix)
        table = doc.add_table(rows=6, cols=6)
        table.style = 'Table Grid'
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Header
        # Empty cell in top left
        table.rows[0].cells[0].text = "Impact \n Likelihood"
        table.rows[0].cells[0].paragraphs[0].runs[0].bold = True

        # Header columns (Impact)
        for j, level in enumerate(levels, 1):
            table.rows[0].cells[j].text = level
            table.rows[0].cells[j].paragraphs[0].runs[0].bold = True
        
        # Header rows (Likelihood) and matrix content
        for i, likelihood in enumerate(levels, 1):
            # Header row
            table.rows[i].cells[0].text = likelihood
            table.rows[i].cells[0].paragraphs[0].runs[0].bold = True

            # Matrix content
            for j, impact in enumerate(levels, 1):
                risk_level = self.RISK_MATRIX.get((likelihood, impact), "")
                table.rows[i].cells[j].text = risk_level

                # Colors the cells based on risk level
                cell = table.rows[i].cells[j]
                if risk_level == "Very High":
                    # Dark Red
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(139, 0, 0)
                elif risk_level == "High":
                    # Red
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(220, 20, 60)
                elif risk_level == "Medium":
                    # Orange
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(255, 140, 0)
                elif risk_level == "Low":
                    # Dark Yellow
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(184, 134, 11)
                elif risk_level == "Very Low":
                    # Green
                    cell.paragraphs[0].runs[0].font.color.rgb = RGBColor(34, 139, 34)

        doc.add_paragraph()  # Space after the table

    def show_help(self):
        """Show help window with criteria descriptions"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Assessment Criteria - Help")
        help_window.geometry("1200x540")
        help_window.configure(bg=self.COLORS['white'])
        help_window.resizable(True, True)
        
        # Center the window
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Title
        title_label = tk.Label(help_window, text="Risk Assessment Criteria Descriptions", 
                              font=('Segoe UI',  16, 'bold'),
                              bg=self.COLORS['white'], fg=self.COLORS['dark'])
        title_label.pack(pady=(20, 15))
        
        # Create scrollable frame for the table
        canvas = tk.Canvas(help_window, bg=self.COLORS['white'], highlightthickness=0)
        scrollbar = tk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create table frame
        table_frame = tk.Frame(scrollable_frame, bg=self.COLORS['white'])
        table_frame.pack(fill='both', expand=True, padx=15, pady=10)
        
        # Table headers
        header_frame = tk.Frame(table_frame, bg=self.COLORS['primary'], relief='ridge', bd=1)
        header_frame.pack(fill='x', pady=(0, 2))
        
        # Configure grid for better column control
        header_frame.grid_columnconfigure(0, weight=0, minsize=250)
        header_frame.grid_columnconfigure(1, weight=1)
        
        criterion_header = tk.Label(header_frame, text="Criterion", font=('Segoe UI', 12, 'bold'),
                                   bg=self.COLORS['primary'], fg=self.COLORS['white'], anchor='w',
                                   padx=15, pady=10)
        criterion_header.grid(row=0, column=0, sticky='ew')
        
        desc_header = tk.Label(header_frame, text="Description", font=('Segoe UI', 12, 'bold'),
                              bg=self.COLORS['primary'], fg=self.COLORS['white'], anchor='w',
                              padx=15, pady=10)
        desc_header.grid(row=0, column=1, sticky='ew')
        
        # Criteria descriptions
        criteria_help = {
            "Vulnerability Level": "Measures the presence and severity of known security vulnerabilities in the system. Lower scores indicate well-patched systems with no known vulnerabilities, while higher scores indicate systems with actively exploitable vulnerabilities.",
            "Access Control": "Evaluates the strength of physical and logical access controls protecting the system. This includes authentication mechanisms, authorization policies, and physical security measures.",
            "Defense Capability": "Assesses the effectiveness of security countermeasures and detection systems. This includes firewalls, intrusion detection, monitoring systems, and incident response capabilities.",
            "Operational Impact": "Measures the potential impact on mission operations if the threat materializes. This considers service disruption, data loss, and effects on critical mission functions.",
            "Recovery Time": "Evaluates the time and resources required to restore normal operations after a security incident. This includes backup systems, recovery procedures, and business continuity planning."
        }
        
        # Add criteria rows
        for i, (criterion, description) in enumerate(criteria_help.items()):
            # Row frame
            row_color = self.COLORS['light'] if i % 2 == 0 else self.COLORS['white']
            row_frame = tk.Frame(table_frame, bg=row_color, relief='ridge', bd=1)
            row_frame.pack(fill='x', pady=1)
            
            # Configure grid for consistent column widths
            row_frame.grid_columnconfigure(0, weight=0, minsize=250)
            row_frame.grid_columnconfigure(1, weight=1)
            
            # Criterion name (left column)
            criterion_label = tk.Label(row_frame, text=criterion, 
                                      font=('Segoe UI', 11, 'bold'),
                                      bg=row_color, fg=self.COLORS['dark'], anchor='nw',
                                      padx=15, pady=8, wraplength=220, justify='left')
            criterion_label.grid(row=0, column=0, sticky='new')
            
            # Description (right column)
            desc_label = tk.Label(row_frame, text=description,
                                 font=('Segoe UI', 11),
                                 bg=row_color, fg='#495057', anchor='nw',
                                 padx=15, pady=8, wraplength=800, justify='left')
            desc_label.grid(row=0, column=1, sticky='new')
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        # Mouse wheel scrolling for help window only
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        # Keep track of bound widgets for cleanup
        bound_widgets = []
        
        # Bind mouse wheel only to the help window and its children
        canvas.bind("<MouseWheel>", _on_mousewheel)
        scrollable_frame.bind("<MouseWheel>", _on_mousewheel)
        bound_widgets.extend([canvas, scrollable_frame])
        
        # Ensure proper cleanup when window is closed
        def on_help_window_close():
            # Remove all mouse wheel bindings
            for widget in bound_widgets:
                try:
                    widget.unbind("<MouseWheel>")
                except:
                    pass  # Widget might be already destroyed
            help_window.destroy()
        
        help_window.protocol("WM_DELETE_WINDOW", on_help_window_close)
        
        # Focus on help window
        help_window.focus_set()
        
    def category_to_value(self, category):
        """Converts a category string to a numeric value (0-1 range)"""
        category_mapping = {
            'Very Low': 0.05,
            'Low': 0.1,
            'Medium': 0.4,
            'High': 0.7,
            'Very High': 0.9
        }
        return category_mapping.get(category, 0.5)
    
    def likelihood_to_base_value(self, likelihood_val):
        """Converts likelihood value to base assessment value (1-5 range)"""
        # Convert from [0,1] to [1,5] range
        # We need to reverse the formula: likelihood = (quadratic_mean - 1) / 4
        # So: quadratic_mean = likelihood * 4 + 1
        quadratic_mean = likelihood_val * 4 + 1
        
        # For simplicity, we'll use the quadratic mean as the base value
        # This assumes all three values (Vulnerability, Access, Defense) are equal
        base_value = max(1, min(5, int(round(quadratic_mean))))
        return base_value
    
    def parse_legacy_word_report(self, doc):
        """Parse a legacy Word report and extract threat data"""
        threats_data = {}
        
        try:
            # First, extract mission type from paragraphs
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                if text.startswith("Mission Type:"):
                    mission_type = text.replace("Mission Type:", "").strip()
                    # Update the dropdown if the mission type is in the list
                    if mission_type in self.MISSION_TYPES:
                        self.mission_type_var.set(mission_type)
                    elif mission_type:
                        # If not in the list, add it temporarily for this session
                        self.mission_type_var.set(mission_type)
            
            # Look for tables in the document
            for table in doc.tables:
                # Check if this is the main threats table
                if len(table.rows) > 1 and len(table.columns) >= 2:
                    # Check header row
                    header_row = table.rows[0]
                    headers = [cell.text.strip() for cell in header_row.cells]
                    
                    # Look for threat/probability table
                    if ('Threat' in headers and 'Probability' in headers) or \
                       ('Threat' in headers and any('Prob' in h for h in headers)):
                        
                        threat_col = -1
                        prob_col = -1
                        
                        # Find column indices
                        for i, header in enumerate(headers):
                            if 'Threat' in header:
                                threat_col = i
                            elif 'Probability' in header or 'Prob' in header:
                                prob_col = i
                        
                        if threat_col >= 0 and prob_col >= 0:
                            # Parse data rows
                            for row in table.rows[1:]:  # Skip header
                                if len(row.cells) > max(threat_col, prob_col):
                                    threat_name = row.cells[threat_col].text.strip()
                                    probability = row.cells[prob_col].text.strip()
                                    
                                    # Map probability values
                                    prob_mapping = {
                                        'Very Low': 'Very Low', 'Low': 'Low', 'Medium': 'Medium',
                                        'High': 'High', 'Very High': 'Very High',
                                        'VL': 'Very Low', 'L': 'Low', 'M': 'Medium',
                                        'H': 'High', 'VH': 'Very High'
                                    }
                                    
                                    probability = prob_mapping.get(probability, probability)
                                    
                                    # Only add if it's a valid threat
                                    if threat_name in self.THREATS and probability in ['Very Low', 'Low', 'Medium', 'High', 'Very High']:
                                        # Extract asset categories from detailed analysis
                                        asset_categories = self.extract_asset_categories_from_doc(doc, threat_name)
                                        
                                        threats_data[threat_name] = {
                                            'likelihood': probability,
                                            'assets': asset_categories
                                        }
            
        except Exception as e:
            print(f"Error parsing legacy Word report: {e}")
        
        return threats_data
    
    def extract_asset_categories_from_doc(self, doc, threat_name):
        """Extract asset categories mentioned for a specific threat"""
        asset_categories = []
        
        try:
            # Look for the threat in the document text
            in_threat_section = False
            
            for paragraph in doc.paragraphs:
                text = paragraph.text.strip()
                
                # Check if we're in the threat section
                if text == threat_name:
                    in_threat_section = True
                    continue
                elif in_threat_section and text in self.THREATS:
                    # We've moved to another threat section
                    break
                elif in_threat_section and text.startswith('Security Controls'):
                    # End of threat section
                    break
                
                # If we're in the threat section, look for asset mentions
                if in_threat_section and text:
                    # Look for asset categories in the text
                    for main_cat, sub_cat in self.ASSET_CATEGORIES:
                        if main_cat.lower() in text.lower() or sub_cat.lower() in text.lower():
                            if main_cat not in asset_categories:
                                asset_categories.append(main_cat)
                            if sub_cat not in asset_categories:
                                asset_categories.append(sub_cat)
        
        except Exception as e:
            print(f"Error extracting asset categories for {threat_name}: {e}")
        
        # If no specific assets found, default to common ones
        if not asset_categories:
            asset_categories = ['Ground', 'Space', 'Link', 'User']
        
        return asset_categories

def main():
    """Main function"""
    root = tk.Tk()
    app = RiskAssessmentTool(root)
    root.mainloop()

if __name__ == "__main__":
    main()
