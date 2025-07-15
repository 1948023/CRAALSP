#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Main Interface - Risk Assessment Tool Suite
Purpose: Central launcher for all risk assessment tools
Author: Thesis work for space program risk assessment tool Giuseppe Nonni 1948023 giuseppe.nonni@gmail.com
"""

import tkinter as tk
from tkinter import ttk, messagebox
import subprocess
import sys
import os
import threading
from datetime import datetime

try:
    from PIL import Image, ImageTk
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

def get_base_path():
    """Get the base path for the application (handles both .py and .exe execution)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
        return os.path.dirname(sys.executable)
    else:
        # Running as Python script
        return os.path.dirname(os.path.abspath(__file__))

class MainInterface:
    """Main interface for the Risk Assessment Tool Suite"""
    
    # Color configuration matching the other tools
    COLORS = {
        'primary': '#4a90c2', 'secondary': '#dc3545', 'success': '#28a745',
        'white': '#ffffff', 'light': '#f8f9fa', 'dark': '#2c3e50',
        'gray': '#6c757d', 'blue': '#007bff', 'green': '#d4edda',
        'yellow': '#fff3cd', 'red': '#f8d7da', 'dark_red': '#dc3545',
        'criteria_header': '#5a67d8', 'criteria_bg': '#edf2f7',
        'asset_header': '#38b2ac', 'asset_bg': '#f0fff4'
    }
    
    # Tools configuration
    TOOLS = [
        {
            'name': 'BID Phase',
            'file': '0-BID.exe',
            'description': 'Calculate risk value of an ITT from project category',
            'color': '#4a90c2',
            'icon': '📊'
        },
        {
            'name': 'Risk Assessment 0-A',
            'file': '1-Risk_Assessment_0-A.exe',
            'description': 'Preliminary risk assessment for space missions',
            'color': '#5a67d8',
            'icon': '🔍'
        },
        {
            'name': 'Risk Assessment',
            'file': '2-Risk_Assessment.exe',
            'description': 'Complete risk assessment tool',
            'color': '#38b2ac',
            'icon': '⚠️'
        },
        {
            'name': 'Attack Graph Analyzer',
            'file': '3-attack_graph_analyzer.exe',
            'description': 'Analyze relationships between threats in space systems',
            'color': '#dc3545',
            'icon': '🔗'
        }
    ]
    
    def __init__(self, root):
        self.root = root
        self.root.title("Risk Assessment Tool Suite")
        self.setup_scaling()
        self.setup_ui()
        self.running_processes = {}
        
    def setup_scaling(self):
        """Calculate scale factors based on screen resolution"""
        screen_area = self.root.winfo_screenwidth() * self.root.winfo_screenheight()
        self.scale_factor = max(0.8, min(2.0, (screen_area / (1920 * 1080)) ** 0.5))
        
        # Scaled dimensions
        self.scaled_font_size = max(10, int(12 * self.scale_factor))
        self.scaled_title_font = max(16, int(20 * self.scale_factor))
        self.scaled_button_font = max(11, int(13 * self.scale_factor))
        self.scaled_padding = max(8, int(10 * self.scale_factor))
        self.scaled_button_padding = max(20, int(25 * self.scale_factor))
        
    def setup_ui(self):
        """Setup main UI structure"""
        # Set window size and center it
        window_width = 1200
        window_height = 800
        
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.configure(bg=self.COLORS['white'])
        self.root.resizable(True, True)
        
        # Header
        self.create_header()
        
        # Main content area
        self.create_main_content()
        
        # Status bar
        self.create_status_bar()
        
    def create_header(self):
        """Create header section"""
        header_frame = tk.Frame(self.root, bg=self.COLORS['primary'], height=120)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)
        
        # Title
        title_label = tk.Label(
            header_frame, 
            text="Risk Assessment Tool Suite",
            font=('Segoe UI', self.scaled_title_font, 'bold'),
            bg=self.COLORS['primary'], 
            fg=self.COLORS['white']
        )
        title_label.pack(pady=(20, 5))
        
        # Subtitle
        subtitle_label = tk.Label(
            header_frame, 
            text="Integrated tool for risk analysis in space missions",
            font=('Segoe UI', self.scaled_font_size, 'italic'),
            bg=self.COLORS['primary'], 
            fg=self.COLORS['white']
        )
        subtitle_label.pack(pady=(0, 10))
        
    def create_main_content(self):
        """Create main content area with tool cards"""
        # Main container with padding
        main_container = tk.Frame(self.root, bg=self.COLORS['white'])
        main_container.pack(fill='both', expand=True, padx=40, pady=30)
        
        # Tools grid
        tools_frame = tk.Frame(main_container, bg=self.COLORS['white'])
        tools_frame.pack(fill='both', expand=True)
        
        # Configure grid
        tools_frame.grid_columnconfigure(0, weight=1)
        tools_frame.grid_columnconfigure(1, weight=1)
        
        # Create tool cards
        for i, tool in enumerate(self.TOOLS):
            row = i // 2
            col = i % 2
            
            self.create_tool_card(tools_frame, tool, row, col)
            
    def create_tool_card(self, parent, tool, row, col):
        """Create a card for each tool"""
        # Card frame
        card_frame = tk.Frame(
            parent, 
            bg=self.COLORS['white'], 
            relief='ridge', 
            bd=2,
            padx=20,
            pady=20
        )
        card_frame.grid(row=row, column=col, padx=15, pady=15, sticky='nsew')
        parent.grid_rowconfigure(row, weight=1)
        
        # Icon and title frame
        header_frame = tk.Frame(card_frame, bg=self.COLORS['white'])
        header_frame.pack(fill='x', pady=(0, 10))
        
        # Icon
        icon_label = tk.Label(
            header_frame,
            text=tool['icon'],
            font=('Segoe UI', self.scaled_title_font + 4),
            bg=self.COLORS['white'],
            fg=tool['color']
        )
        icon_label.pack(side='left')
        
        # Title
        title_label = tk.Label(
            header_frame,
            text=tool['name'],
            font=('Segoe UI', self.scaled_font_size + 2, 'bold'),
            bg=self.COLORS['white'],
            fg=self.COLORS['dark']
        )
        title_label.pack(side='left', padx=(10, 0))
        
        # Description
        desc_label = tk.Label(
            card_frame,
            text=tool['description'],
            font=('Segoe UI', self.scaled_font_size),
            bg=self.COLORS['white'],
            fg=self.COLORS['gray'],
            wraplength=280,
            justify='left'
        )
        desc_label.pack(fill='x', pady=(0, 15))
        
        # Button frame
        button_frame = tk.Frame(card_frame, bg=self.COLORS['white'])
        button_frame.pack(fill='x')
        
        # Run button
        run_button = tk.Button(
            button_frame,
            text="Run",
            font=('Segoe UI', self.scaled_button_font, 'bold'),
            bg=tool['color'],
            fg=self.COLORS['white'],
            relief='flat',
            padx=self.scaled_button_padding,
            pady=8,
            cursor='hand2',
            command=lambda t=tool: self.run_tool(t)
        )
        run_button.pack(side='left')
        
        # Add hover effects
        run_button.bind('<Enter>', lambda e, btn=run_button, color=tool['color']: 
                       self.on_button_hover(btn, color))
        run_button.bind('<Leave>', lambda e, btn=run_button, color=tool['color']: 
                       self.on_button_leave(btn, color))
        
        # Status label for this tool
        status_label = tk.Label(
            button_frame,
            text="",
            font=('Segoe UI', self.scaled_font_size - 1),
            bg=self.COLORS['white'],
            fg=self.COLORS['gray']
        )
        status_label.pack(side='right', padx=(10, 0))
        
        # Store reference to status label
        tool['status_label'] = status_label
        
    def on_button_hover(self, button, color):
        """Handle button hover effect"""
        # Darken the color slightly
        button.configure(bg=self.darken_color(color))
        
    def on_button_leave(self, button, color):
        """Handle button leave effect"""
        button.configure(bg=color)
        
    def darken_color(self, color):
        """Darken a hex color by 15%"""
        # Remove # if present
        color = color.lstrip('#')
        
        # Convert to RGB
        r = int(color[0:2], 16)
        g = int(color[2:4], 16)
        b = int(color[4:6], 16)
        
        # Darken by 15%
        r = max(0, int(r * 0.85))
        g = max(0, int(g * 0.85))
        b = max(0, int(b * 0.85))
        
        # Convert back to hex
        return f"#{r:02x}{g:02x}{b:02x}"
        
    def create_status_bar(self):
        """Create status bar at bottom"""
        self.status_bar = tk.Frame(self.root, bg=self.COLORS['light'], height=30)
        self.status_bar.pack(fill='x', side='bottom')
        self.status_bar.pack_propagate(False)
        
        self.status_label = tk.Label(
            self.status_bar,
            text="Ready",
            font=('Segoe UI', self.scaled_font_size - 1),
            bg=self.COLORS['light'],
            fg=self.COLORS['dark']
        )
        self.status_label.pack(side='left', padx=10, pady=5)
        
        # Time label
        self.time_label = tk.Label(
            self.status_bar,
            text="",
            font=('Segoe UI', self.scaled_font_size - 1),
            bg=self.COLORS['light'],
            fg=self.COLORS['gray']
        )
        self.time_label.pack(side='right', padx=10, pady=5)
        
        # Update time
        self.update_time()
        
    def update_time(self):
        """Update time display"""
        current_time = datetime.now().strftime("%H:%M:%S")
        self.time_label.config(text=current_time)
        self.root.after(1000, self.update_time)
        
    def run_tool(self, tool):
        """Run the selected tool"""
        if tool['file'] in self.running_processes:
            messagebox.showwarning(
                "Warning", 
                f"The tool {tool['name']} is already running."
            )
            return
        
        # Get the full path to the executable
        base_path = get_base_path()
        exe_path = os.path.join(base_path, tool['file'])
        
        # Check if executable exists
        if not os.path.exists(exe_path):
            messagebox.showerror(
                "Error", 
                f"The executable {tool['file']} was not found in {base_path}."
            )
            return
            
        # Update status
        self.update_status(f"Starting {tool['name']}...")
        tool['status_label'].config(text="Starting...", fg=self.COLORS['blue'])
        
        # Run in separate thread
        thread = threading.Thread(target=self.execute_tool, args=(tool, exe_path))
        thread.daemon = True
        thread.start()
        
    def execute_tool(self, tool, exe_path):
        """Execute tool in separate thread"""
        try:
            # Execute the .exe file directly
            process = subprocess.Popen(
                [exe_path], 
                stdout=subprocess.PIPE, 
                stderr=subprocess.PIPE,
                text=True,
                creationflags=subprocess.CREATE_NEW_CONSOLE if os.name == 'nt' else 0
            )
            
            # Store process reference
            self.running_processes[tool['file']] = process
            
            # Update UI in main thread
            self.root.after(0, lambda: tool['status_label'].config(text="Running...", fg=self.COLORS['success']))
            self.root.after(0, lambda: self.update_status(f"{tool['name']} is running"))
            
            # Wait for process to complete
            stdout, stderr = process.communicate()
            
            # Process completed
            if process.returncode == 0:
                self.root.after(0, lambda: tool['status_label'].config(text="Completed", fg=self.COLORS['success']))
                self.root.after(0, lambda: self.update_status(f"{tool['name']} completed successfully"))
            else:
                self.root.after(0, lambda: tool['status_label'].config(text="Error", fg=self.COLORS['secondary']))
                self.root.after(0, lambda: self.update_status(f"{tool['name']} completed with errors"))
                
        except Exception as e:
            # Error occurred
            error_msg = f"Error running {tool['name']}: {str(e)}"
            self.root.after(0, lambda: tool['status_label'].config(text="Error", fg=self.COLORS['secondary']))
            self.root.after(0, lambda: self.update_status(error_msg))
            
        finally:
            # Remove from running processes
            if tool['file'] in self.running_processes:
                del self.running_processes[tool['file']]
            
            # Clear status after 3 seconds
            self.root.after(3000, lambda: tool['status_label'].config(text="", fg=self.COLORS['gray']))
            
    def update_tool_status(self, tool, status, color):
        """Update tool status in UI"""
        tool['status_label'].config(text=status, fg=color)
        
        # Clear status after a few seconds if completed
        if status in ["Completed", "Error"]:
            self.root.after(3000, lambda: tool['status_label'].config(text=""))
            
    def update_status(self, message):
        """Update status bar message"""
        self.status_label.config(text=message)
        
    def on_closing(self):
        """Handle window closing"""
        if self.running_processes:
            if messagebox.askokcancel(
                "Closing", 
                "There are processes running. Do you want to close anyway?"
            ):
                # Terminate all running processes
                for process in self.running_processes.values():
                    try:
                        process.terminate()
                    except:
                        pass
                self.root.destroy()
        else:
            self.root.destroy()

def main():
    """Main function"""
    root = tk.Tk()
    app = MainInterface(root)
    
    # Handle window closing
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    # Start the application
    root.mainloop()

if __name__ == "__main__":
    main()