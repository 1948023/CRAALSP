#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Risk Assessment Tool - Clean Version
Optimized and refactored for clarity and maintainability
"""

# Aggiungi questa linea all'inizio del file, dopo gli altri import
from export_import_functions import ExportImportManager

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
import os
import sys
import math
import datetime
import logging

def get_base_path():
    """Get the base path for the application (works with both .py and .exe)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable - look for CSV files next to the .exe
        return os.path.dirname(sys.executable)
    else:
        # Running as script
        return os.path.dirname(os.path.abspath(__file__))

def get_output_path():
    """Get the path where output files should be saved"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable - save next to the .exe
        return os.path.dirname(sys.executable)
    else:
        # Running as script - save in script directory
        return os.path.dirname(os.path.abspath(__file__))

# Import for Word export/import
try:
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Setup logging
logging.basicConfig(level=logging.INFO)

class RiskAssessmentTool:
    """Risk Assessment Tool - Clean and Optimized Version"""
    
    # Color configuration
    COLORS = {
        'primary': '#4a90c2', 'success': '#28a745', 'white': '#ffffff',
        'light': '#f8f9fa', 'dark': '#2c3e50', 'gray': '#6c757d',
        'criteria_header': '#5a67d8', 'criteria_bg': '#edf2f7'
    }
    
    # Criteria colors - light and attenuated colors to distinguish criteria
    CRITERIA_COLORS = [
        '#ffeaa7',  # Light yellow - Criterion 1
        '#dda0dd',  # Light lilac - Criterion 2  
        '#98fb98',  # Light mint - Criterion 3
        '#f0e68c',  # Light khaki - Criterion 4
        '#ffd1dc',  # Pale pink - Criterion 5
        '#e0ffff',  # Ice blue - Criterion 6
        '#ffe4e1',  # Misty rose - Criterion 7
        '#f5deb3',  # Wheat - Criterion 8
        '#d3d3d3'   # Light gray - Criterion 9
    ]
    
    # Threat criteria (7 criteria: 5 for likelihood, 2 for impact) - Transposed format
    CRITERIA_DATA_THREAT = [
        ["Score", "Vulnerability effectiveness", "Mitigation Presence", "Detection Probability", "Access Complexity", "Privilege Requirement", "Response Delay", "Resilience Impact"],
        ["Score 1 (Very Low)", "No known or already resolved vulnerabilities", "Multi-level countermeasures in place and validated", "Real-time, centralized, and automated detection", "Access strongly protected by physical/logical measures", "Requires root/admin access", "Predefined automated response", "No disruption: Full operability with local redundancies, automatic failover, and tested continuity plans"],
        ["Score 2 (Low)", "Known vulnerability, mitigated through hardening and patches", "Robust countermeasures but not regularly tested", "Automated but not centralized detection", "Moderately protected access (VPN, ACL, bastion host)", "Elevated privileges but not root", "Quick response thanks to well-defined procedures", "Temporary impact: Quick restoration via documented, semi-automated procedures. No lasting degradation"],
        ["Score 3 (Moderate)", "Known vulnerability, but only partially mitigated", "Limited or isolated countermeasures", "Manual or retrospective detection only", "Access protected with weaker controls", "Standard user privileges", "Manual but formalized response", "Partial degradation: Minimum operational capacity maintained. Manual intervention and noticeable recovery time required"],
        ["Score 4 (High)", "Known vulnerability, with no effective mitigations", "Weak or outdated countermeasures", "Occasional or incorrect detection", "Access easily accessible by remote attackers", "Minimal privileges or no authentication", "Slow or poorly coordinated response", "Severe impact: Critical unavailability. Recoverable only with urgent external intervention"],
        ["Score 5 (Very High)", "Actively exploitable vulnerability, with no defenses", "No relevant countermeasures", "No detection capability", "Completely open or physically accessible access", "No privileges required", "No response capability", "Irreversible loss: Asset permanently disabled or destroyed. No recovery possible"]
    ]
    
    # Asset criteria (9 criteria: 4 for likelihood, 5 for impact) - Transposed format
    CRITERIA_DATA_ASSET = [
        ["Score", "Dependency", "Penetration", "Cyber Maturity", "Trust", "Performance", "Schedule", "Costs", "Reputation", "Recovery"],
        ["Score 1 (Very Low)", "Asset not involved in mission-critical functions", "No access or isolated user-level access", "Mature, audited, and mission-integrated cyber governance system with real-time threat management", "Strategic partner under strict control, with shared security responsibility and continuous assurance", "Minimal or no impact", "Minimal or no impact", "Minimal or no impact", "Issue contained internally with no external reputational impact", "Limited damage to the mission. Up to 1 month to resumption of normal commercial operations"],
        ["Score 2 (Low)", "Useful support asset ", "User-level access to general ground segment components", "Integrated and proactive cybersecurity program; includes vulnerability management and incident drills", "Stakeholder trusted, with contractual obligations and validated controls", "Moderate reduction, Some approach retained", "Additional activities required, able to meet need dates", "Cost increase < 5%", "Slight reputational damage; disclosure required to customers and reassurance efforts toward external stakeholders", "Minor damage to the mission  resulting in up to 3 months to resumption of normal commercial operations"],
        ["Score 3 (Moderate)", "Relationship important for multiple business processes", "Admin-level access to mission services", "Organization enforces a cybersecurity policy with partially proactive security practices", "Stakeholder known and generally aligned. Moderate assurance level", "Moderate reduction, but workarounds available", "Project team milestone slip <= 1 month", "Cost increase > 5%", "Noticeable reputational harm; loss of customer trust, media coverage, and regulatory disclosure required", "Moderate damage to the mission  resulting in up to 6 months to resumption of normal commercial operations"],
        ["Score 4 (High)", "Asset supporting several mission services", "Admin access to mission-critical components", "Security rules exist but are scattered. Limited integration with mission security architecture", "Stakeholder considered low-risk but no formal guarantees", "Major reduction, but workarounds available", "Project milestone slip >= 1 month or project critical path impacted", "Cost increase > 10%", "Serious reputational damage; loss of investor confidence, negative media exposure, and client disengagement", "Significant damage to the mission  resulting in up to 1 year to resumption of normal commercial operations"],
        ["Score 5 (Very High)", "Essential asset", "Full privileged access to core mission infrastructure", "Minimal cybersecurity procedures. No defined response to cyber incidents", "No trust relationship; stakeholder identity or intent unknown", "Unacceptable, no alternatives exist", "Can't achieve major project milestone", "Cost increase > 15%", "Irreparable reputational harm; international fallout, industry-wide loss of credibility, potential business closure", "Catastrophic damage long term (more than  1 year) or complete loss of mission  indefinitely"]
    ]
    
    # Risk matrix (ISO 27005)
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
        ("Very Low", "Very Low"): "Very Low"
    }

    def __init__(self, root):
        self.root = root
        self.root.title("Risk Assessment Tool - Clean Version")
        self.root.state('zoomed')
        self.root.configure(bg=self.COLORS['white'])
          # Data storage - separated for threats and assets
        self.threat_data = {}  # Saved data for threats
        self.asset_data = {}   # Saved data for assets
        
        # Threat window variables
        self.threat_combo_vars = {}   # ComboBox variables for threat window
        self.threat_impact_entries = {}  # Table widgets for threat window
        
        # Asset window variables  
        self.asset_combo_vars = {}   # ComboBox variables for asset window
        self.asset_impact_entries = {}  # Table widgets for asset window
        
        # Current window variables (will be set when windows are opened)
        self.combo_vars = {}       # ComboBox variables 
        self.impact_entries = {}   # Table widgets
        self.selected_threat_var = tk.StringVar()  # Initialize immediately to avoid None errors
        self.selected_asset_var = tk.StringVar()   # Initialize for asset selection
        
        # Asset assessment dictionary indexed by asset
        self.asset_assessment_dict = {}  # Dictionary indexed by asset for inference
        
        # Initialize export/import manager
        self.export_import_manager = ExportImportManager(self)
        
        # ========== CONTROLS STATE MANAGEMENT ==========
        # Track applied controls (simplified approach)
        self.applied_controls = []  # List of applied control IDs
        
        # Mission info
        self.mission_name = ""
        self.mission_type = ""

        # Load external data
        self.load_threats_from_csv()
        # Load assets from CSV
        self.load_assets_from_csv()
        # Load controls from CSV
        self.load_controls_from_csv()
        
        # Setup custom styles
        self.setup_combobox_styles()
        
        # Create interface
        self.create_interface()
        
        # Initialize button states
        self.update_threat_analysis_button_state()
        
        # Set up close confirmation
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def on_closing(self):
        """Handle window closing with confirmation dialog"""
        result = messagebox.askyesno(
            "Confirm Exit",
            "Are you sure you want to exit?\nNot exported values will be lost.",
            icon='warning'
        )
        if result:
            self.root.destroy()

    # ========== CONTROLS STATE MANAGEMENT (SIMPLIFIED) ==========
    
    def is_control_applied(self, control_id):
        """Check if a control is already applied"""
        return control_id in self.applied_controls
    
    def apply_control(self, control_id, control_data):
        """Apply a single control and modify threat criteria directly"""
        if self.is_control_applied(control_id):
            print(f"[CONTROLS] Control {control_id} already applied")
            return False
        
        # Apply control effects directly to threat data
        self.apply_control_effects(control_data)
        
        # Add to applied list
        self.applied_controls.append(control_id)
        
        # Update button state
        self.update_threat_analysis_button_state()
        
        print(f"[CONTROLS] Applied control {control_id}: {control_data['title']}")
        return True
    
    def remove_control(self, control_id):
        """Remove a control and reverse its effects"""
        if not self.is_control_applied(control_id):
            print(f"[CONTROLS] Control {control_id} not applied")
            return False
        
        # Get control data to reverse effects
        control_data = self.get_control_by_id(control_id)
        if control_data:
            self.reverse_control_effects(control_data)
        
        # Remove from applied list
        self.applied_controls.remove(control_id)
        
        # Update button state
        self.update_threat_analysis_button_state()
        
        print(f"[CONTROLS] Removed control {control_id}")
        return True
    
    def get_control_by_id(self, control_id):
        """Get control data by ID"""
        for control in self.CONTROLS:
            if control['id'] == control_id:
                return control
        return None
    
    def find_matching_threats(self, threat_name):
        """Find threats that match the given threat name using flexible matching"""
        matching_threats = []
        threat_name_lower = threat_name.lower().strip()
        
        # First try exact match
        if threat_name in self.threat_data:
            matching_threats.append(threat_name)
            return matching_threats
        
        # Then try flexible matching with all threats in the dataset
        for actual_threat in self.threat_data.keys():
            actual_lower = actual_threat.lower()
            
            # Check if names contain each other
            if (threat_name_lower in actual_lower or 
                actual_lower in threat_name_lower or
                self.normalize_threat_name(threat_name_lower) == self.normalize_threat_name(actual_lower)):
                matching_threats.append(actual_threat)
        
        # If no matches found, try against the THREATS list
        if not matching_threats:
            for threat in self.THREATS:
                threat_lower = threat.lower()
                if (threat_name_lower in threat_lower or 
                    threat_lower in threat_name_lower or
                    self.normalize_threat_name(threat_name_lower) == self.normalize_threat_name(threat_lower)):
                    if threat in self.threat_data:
                        matching_threats.append(threat)
        
        return matching_threats
    
    def get_assets_for_segment(self, control_segment):
        """Get asset keys that match the control segment (supports multiple segments separated by comma)"""
        if not control_segment:
            return []
        
        # Split multiple segments separated by comma and normalize each
        segments = [s.strip().lower() for s in control_segment.split(',') if s.strip()]
        
        # Special case mappings for each segment
        normalized_segments = []
        for segment in segments:
            if segment == 'human':
                normalized_segments.append('user ground segment')
            else:
                normalized_segments.append(segment)
        
        matching_asset_keys = []
        
        # Check each asset category
        for i, (category, subcategory, asset) in enumerate(self.ASSET_CATEGORIES):
            category_lower = category.lower()
            subcategory_lower = subcategory.lower()
            
            # Check if this asset matches any of the segments
            match_found = False
            
            for segment_lower in normalized_segments:
                # 1. Direct category match (Ground, Space, Link, User)
                if segment_lower == category_lower:
                    match_found = True
                    break
                
                # 2. Subcategory match
                elif segment_lower in subcategory_lower or subcategory_lower == segment_lower:
                    match_found = True
                    break
                
                # 3. Special case mappings
                elif (segment_lower == 'user ground segment' and 
                      category_lower == 'ground' and 'user' in subcategory_lower):
                    match_found = True
                    break
                
                elif (segment_lower == 'human resources' and 
                      category_lower == 'ground' and 'user' in subcategory_lower):
                    match_found = True
                    break
            
            if match_found:
                # Add the asset key in the format used by the threat assessment
                # Format: "{i+1}_probability" for threat assessment
                asset_key_threat = f"{i+1}_probability"
                matching_asset_keys.append(asset_key_threat)
                
                # Also add the index-based key for other potential uses
                matching_asset_keys.append(str(i))
                
                # Add full description format if different systems use it
                full_description = f"{category}.{subcategory}.{asset}"
                matching_asset_keys.append(full_description)
        
        print(f"[CONTROLS] Segment '{control_segment}' (parsed as: {segments}) matches asset keys: {matching_asset_keys}")
        return matching_asset_keys
    
    def apply_control_effects(self, control_data):
        """Apply control effects directly to threat criteria for compatible assets only"""
        # Parse threats and criteria from control
        threats_addressed = [t.strip() for t in control_data.get('threats_addressed', '').split(',') if t.strip()]
        criteria_list = [c.strip() for c in control_data.get('criteria', '').split(',') if c.strip()]
        control_segment = control_data.get('segment', '').strip()
        
        # Get compatible asset keys for this control's segment
        compatible_assets = self.get_assets_for_segment(control_segment)
        
        if not compatible_assets:
            print(f"[CONTROLS] Warning: No compatible assets found for segment '{control_segment}'")
            return
        
        # Map criteria names to threat table column indices (0-based for actual data storage)
        criteria_mapping = {
            'vulnerability effectiveness': 0,  # Column 0 - Likelihood
            'vulnerability': 0,
            'mitigation presence': 1,          # Column 1 - Likelihood  
            'mitigation': 1,
            'detection probability': 2,        # Column 2 - Likelihood
            'detection': 2,
            'access complexity': 3,            # Column 3 - Likelihood
            'access': 3,
            'privilege requirement': 4,        # Column 4 - Likelihood
            'privilege': 4,
            'response delay': 5,               # Column 5 - Impact
            'response': 5,
            'resilience impact': 6,            # Column 6 - Impact
            'resilience': 6
        }
        
        # For each threat this control addresses
        for threat_name in threats_addressed:
            # Find matching threats using flexible matching
            matching_threats = self.find_matching_threats(threat_name)
            
            for actual_threat in matching_threats:
                if actual_threat in self.threat_data:
                    threat_info = self.threat_data[actual_threat]
                    
                    # For each criteria this control improves
                    for criteria_name in criteria_list:
                        criteria_lower = criteria_name.lower()
                        
                        # Find matching criteria mapping
                        for key, column_index in criteria_mapping.items():
                            if key in criteria_lower:
                                # Apply only to compatible assets for this threat
                                for asset_key, asset_scores in threat_info.items():
                                    # Check if this asset is compatible with the control's segment
                                    if asset_key in compatible_assets:
                                        if isinstance(asset_scores, dict) and str(column_index) in asset_scores:
                                            try:
                                                # Improve the score (decrease by 1, minimum 1)
                                                current_score = int(asset_scores[str(column_index)])
                                                new_score = max(1, current_score - 1)
                                                asset_scores[str(column_index)] = str(new_score)
                                                
                                                print(f"[CONTROLS]  Improved {actual_threat}.{asset_key}[{column_index}]: {current_score} -> {new_score} (segment: {control_segment})")
                                            except ValueError:
                                                print(f"[CONTROLS] Warning: Invalid score format for {actual_threat}.{asset_key}[{column_index}]")
                                        else:
                                            print(f"[CONTROLS] Warning: {actual_threat}.{asset_key}[{column_index}] not found")
                                    else:
                                        print(f"[CONTROLS]  Skipping {actual_threat}.{asset_key} - incompatible with segment '{control_segment}' (compatible: {compatible_assets[:3]}...)")
                                # Continue to check other criteria mappings for this criteria_name
    
    def reverse_control_effects(self, control_data):
        """Reverse control effects by increasing criteria scores for compatible assets only"""
        # Parse threats and criteria from control
        threats_addressed = [t.strip() for t in control_data.get('threats_addressed', '').split(',') if t.strip()]
        criteria_list = [c.strip() for c in control_data.get('criteria', '').split(',') if c.strip()]
        control_segment = control_data.get('segment', '').strip()
        
        # Get compatible asset keys for this control's segment
        compatible_assets = self.get_assets_for_segment(control_segment)
        
        if not compatible_assets:
            print(f"[CONTROLS] Warning: No compatible assets found for segment '{control_segment}' during reversal")
            return
        
        # Map criteria names to threat table column indices (0-based for actual data storage)
        criteria_mapping = {
            'vulnerability effectiveness': 0,  # Column 0 - Likelihood
            'vulnerability': 0,
            'mitigation presence': 1,          # Column 1 - Likelihood  
            'mitigation': 1,
            'detection probability': 2,        # Column 2 - Likelihood
            'detection': 2,
            'access complexity': 3,            # Column 3 - Likelihood
            'access': 3,
            'privilege requirement': 4,        # Column 4 - Likelihood
            'privilege': 4,
            'response delay': 5,               # Column 5 - Impact
            'response': 5,
            'resilience impact': 6,            # Column 6 - Impact
            'resilience': 6
        }
        
        # For each threat this control addressed
        for threat_name in threats_addressed:
            # Find matching threats using flexible matching
            matching_threats = self.find_matching_threats(threat_name)
            
            for actual_threat in matching_threats:
                if actual_threat in self.threat_data:
                    threat_info = self.threat_data[actual_threat]
                    
                    # For each criteria this control improved
                    for criteria_name in criteria_list:
                        criteria_lower = criteria_name.lower()
                        
                        # Find matching criteria mapping  
                        for key, column_index in criteria_mapping.items():
                            if key in criteria_lower:
                                # Reverse effects only for compatible assets for this threat
                                for asset_key, asset_scores in threat_info.items():
                                    # Check if this asset is compatible with the control's segment
                                    if asset_key in compatible_assets:
                                        if isinstance(asset_scores, dict) and str(column_index) in asset_scores:
                                            try:
                                                # Reverse the improvement (increase by 1, maximum 5)
                                                current_score = int(asset_scores[str(column_index)])
                                                new_score = min(5, current_score + 1)
                                                asset_scores[str(column_index)] = str(new_score)
                                                
                                                print(f"[CONTROLS]  Reversed {actual_threat}.{asset_key}[{column_index}]: {current_score} -> {new_score} (segment: {control_segment})")
                                            except ValueError:
                                                print(f"[CONTROLS] Warning: Invalid score format for reversal {actual_threat}.{asset_key}[{column_index}]")
                                        else:
                                            print(f"[CONTROLS] Warning: {actual_threat}.{asset_key}[{column_index}] not found for reversal")
                                    else:
                                        print(f"[CONTROLS]  Skipping reversal {actual_threat}.{asset_key} - incompatible with segment '{control_segment}' (compatible: {compatible_assets[:3]}...)")
                                # Continue to check other criteria mappings for this criteria_name
    
    def clear_all_controls(self):
        """Clear all applied controls by reversing their effects"""
        if not self.applied_controls:
            messagebox.showinfo("No Controls", "No controls are currently applied.")
            return
        
        result = messagebox.askyesno(
            "Clear All Controls",
            f"Remove all {len(self.applied_controls)} applied controls?",
            icon='warning'
        )
        
        if result:
            controls_count = len(self.applied_controls)
            
            # Reverse all applied controls
            for control_id in self.applied_controls.copy():
                self.remove_control(control_id)
            
            # Update main table
            self.update_all_threats_in_main_table()
            
            # Update button state
            self.update_threat_analysis_button_state()
            
            messagebox.showinfo("Success", f"All {controls_count} controls cleared!")
            print(f"[CONTROLS] All {controls_count} controls cleared")

    def clear_all_controls_gui(self):
        """GUI method to clear all controls and update interface"""
        if not hasattr(self, 'applied_controls') or not self.applied_controls:
            messagebox.showinfo("No Controls", "No controls are currently applied.")
            return
        
        # Use the core clear method
        self.clear_all_controls()
        
        # Update controls window if open
        if hasattr(self, 'controls_window') and self.controls_window.winfo_exists():
            try:
                # Refresh controls display to remove visual indicators
                if hasattr(self, 'all_controls_data'):
                    self.display_controls(self.all_controls_data)
                    
                # Clear selected controls checkboxes
                if hasattr(self, 'selected_controls'):
                    for control_id in list(self.selected_controls.keys()):
                        self.selected_controls[control_id].set(False)
                
                # Update impact analysis display
                self.update_impact_analysis()
                
            except Exception as e:
                print(f"[CONTROLS] Warning: Error updating GUI after clear: {e}")
        
        # Close the controls window after clearing
        if hasattr(self, 'controls_window') and self.controls_window.winfo_exists():
            try:
                print(f"[CONTROLS] Closing controls window after clearing all controls")
                self.controls_window.destroy()
            except Exception as e:
                print(f"[CONTROLS] Warning: Error closing controls window: {e}")

    def disable_mousewheel_on_combobox(self, combo):
        """Intelligently handle mouse wheel on combobox to prevent accidental value changes while allowing scroll"""
        def on_mousewheel(event):
            # Check if the combobox dropdown is open
            try:
                if combo.tk.call('ttk::combobox::PopdownIsVisible', combo):
                    # If dropdown is open, allow normal combobox behavior
                    return
                else:
                    # If dropdown is closed, prevent value changes but allow window scrolling
                    # Find the parent canvas to continue scrolling
                    widget = event.widget
                    # Walk up the widget hierarchy to find a canvas
                    while widget:
                        if isinstance(widget, tk.Canvas):
                            widget.yview_scroll(int(-1*(event.delta/120)), "units")
                            break
                        widget = widget.master
                    return "break"  # Prevent combobox value change
            except:
                # Fallback: prevent value changes but allow window scrolling
                widget = event.widget
                while widget:
                    if isinstance(widget, tk.Canvas):
                        widget.yview_scroll(int(-1*(event.delta/120)), "units")
                        break
                    widget = widget.master
                return "break"  # Prevent combobox value change
        
        combo.bind("<MouseWheel>", on_mousewheel)

    def disable_mousewheel_on_checkbox(self, checkbox):
        """Disable mouse wheel on checkbox to prevent accidental changes while allowing scroll"""
        def on_checkbox_mousewheel(event):
            # Don't change checkbox state, but allow scrolling by finding parent canvas
            widget = event.widget
            # Walk up the widget hierarchy to find a canvas
            while widget:
                if isinstance(widget, tk.Canvas):
                    widget.yview_scroll(int(-1*(event.delta/120)), "units")
                    break
                widget = widget.master
            return "break"  # Prevent checkbox value change
        
        checkbox.bind("<MouseWheel>", on_checkbox_mousewheel)

    def setup_impact_mousewheel(self, widget, canvas):
        """Setup mouse wheel scrolling specifically for the impact analysis section"""
        def on_impact_mousewheel(event):
            # Only scroll if the mouse is over the impact section
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                return "break"  # Prevent event propagation
            except:
                pass
        
        # Store the canvas and scroll function for dynamic content updates
        self.impact_canvas = canvas
        self.impact_scroll_function = on_impact_mousewheel
        
        # Bind to the widget and all its children recursively but ONLY for impact
        def bind_impact_mousewheel_recursive(w):
            try:
                # Skip checkbox and combobox widgets to prevent accidental value changes
                if not isinstance(w, (tk.Checkbutton, ttk.Combobox)):
                    w.bind("<MouseWheel>", on_impact_mousewheel)
                
                for child in w.winfo_children():
                    bind_impact_mousewheel_recursive(child)
            except:
                pass
        
        # Start recursive binding for impact only
        bind_impact_mousewheel_recursive(widget)
        
        # Also bind to canvas for safety
        try:
            canvas.bind("<MouseWheel>", on_impact_mousewheel)
        except:
            pass
    
    def apply_impact_mousewheel_to_new_content(self, widget):
        """Apply mouse wheel scrolling to newly created impact content"""
        if not hasattr(self, 'impact_scroll_function'):
            return
            
        def bind_recursive(w):
            try:
                # Skip checkbox and combobox widgets to prevent accidental value changes
                if not isinstance(w, (tk.Checkbutton, ttk.Combobox)):
                    w.bind("<MouseWheel>", self.impact_scroll_function)
                
                for child in w.winfo_children():
                    bind_recursive(child)
            except:
                pass
        
        bind_recursive(widget)

    def setup_controls_mousewheel(self, widget, canvas):
        """Setup mouse wheel scrolling specifically for the controls section"""
        def on_controls_mousewheel(event):
            # Only scroll if the mouse is over the controls section
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
                return "break"  # Prevent event propagation
            except:
                pass
        
        # Bind to the widget and all its children recursively but ONLY for controls
        def bind_controls_mousewheel_recursive(w):
            try:
                # Skip checkbox and combobox widgets to prevent accidental value changes
                if not isinstance(w, (tk.Checkbutton, ttk.Combobox)):
                    w.bind("<MouseWheel>", on_controls_mousewheel)
                
                for child in w.winfo_children():
                    bind_controls_mousewheel_recursive(child)
            except:
                pass
        
        # Start recursive binding for controls only
        bind_controls_mousewheel_recursive(widget)
        
        # Also bind to canvas for safety
        try:
            canvas.bind("<MouseWheel>", on_controls_mousewheel)
        except:
            pass

    def setup_global_mousewheel(self, widget, canvas):
        """Setup global mouse wheel scrolling for any widget relative to a canvas"""
        def on_global_mousewheel(event):
            # Scroll the canvas when mouse wheel is used anywhere in the widget
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass  # Ignore errors if canvas is not scrollable
        
        # Bind to the widget and all its children recursively
        def bind_mousewheel_recursive(w):
            try:
                # Skip checkbox and combobox widgets to prevent accidental value changes
                if not isinstance(w, (tk.Checkbutton, ttk.Combobox)):
                    w.bind("<MouseWheel>", on_global_mousewheel)
                
                # Recursively bind to all children
                for child in w.winfo_children():
                    bind_mousewheel_recursive(child)
            except:
                pass  # Ignore any binding errors
        
        # Start recursive binding
        bind_mousewheel_recursive(widget)
        
        # Also bind directly to the main widget and canvas for safety
        try:
            widget.bind("<MouseWheel>", on_global_mousewheel)
            canvas.bind("<MouseWheel>", on_global_mousewheel)
        except:
            pass
        
        # Add a global catch-all binding to the parent window if it exists
        try:
            parent_window = widget.winfo_toplevel()
            if parent_window:
                parent_window.bind("<MouseWheel>", on_global_mousewheel)
        except:
            pass

    def ensure_mousewheel_on_table_cells(self, canvas):
        """Ensure all threat table cells have mouse wheel scrolling"""
        def on_cell_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except:
                pass
        
        # Apply to all threat cells
        for threat, cells in self.threat_cells.items():
            for cell_type, cell in cells.items():
                cell.bind("<MouseWheel>", on_cell_mousewheel)

    def load_threats_from_csv(self):
        """Load threats from Threat.csv"""
        threats_file = os.path.join(get_base_path(), "Threat.csv")
        self.THREATS = []
        
        try:
            with open(threats_file, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                for row in reader:
                    threat_name = row.get('THREAT', '').strip()
                    if threat_name:
                        self.THREATS.append(threat_name)
            self.THREATS.sort()
            #print(f" Loaded {len(self.THREATS)} threats from {threats_file}")
            
        except FileNotFoundError:
            #print(f" File not found: {threats_file}")
            # Fallback threats
            self.THREATS = [
                "Data Corruption", "Physical/Logical Attack", "Interception/Eavesdropping",
                "Jamming", "Denial-of-Service", "Masquerade/Spoofing", "Replay",
                "Software Threats", "Unauthorized Access/Hijacking", 
                "Tainted hardware components", "Supply Chain"
            ]
        except Exception as e:
            #print(f" Error loading threats: {e}")
            self.THREATS = []

    def load_assets_from_csv(self):
        """Load assets from Asset.csv"""
        assets_file = os.path.join(get_base_path(), "Asset.csv")
        self.ASSET_CATEGORIES = []
        
        try:
            with open(assets_file, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                for row in reader:
                    category = row.get('categories', '').strip()
                    subcategory = row.get('subCategories', '').strip()
                    asset = row.get('asset', '').strip()
                    
                    if category and subcategory and asset:
                        self.ASSET_CATEGORIES.append((category, subcategory, asset))
            
            print(f"[OK] Loaded {len(self.ASSET_CATEGORIES)} asset categories from {assets_file}")
            
        except FileNotFoundError:
            print(f"[ERROR] File not found: {assets_file}")
            # Fallback assets
            self.ASSET_CATEGORIES = [
                ("Ground", "Ground Stations", "Tracking"), ("Ground", "Ground Stations", "Ranging"),
                ("Ground", "Mission Control", "Telemetry processing"), ("Ground", "Mission Control", "Commanding"),
                ("Ground", "Data Processing Centers", "Mission Analysis"), ("Ground", "Remote Terminals", "Network access"),
                ("Ground", "User Ground Segment", "Development"), ("Space", "Platform", "Bus"),
                ("Space", "Payload", "Payload Data Handling Systems"), ("Link", "Link", "Between Platform and Payload"),
                ("User", "User", "Transmission")
            ]
        except Exception as e:
            print(f"[ERROR] Error loading assets: {e}")
            # Fallback assets
            self.ASSET_CATEGORIES = [
                ("Ground", "Ground Stations", "Tracking"), ("Ground", "Ground Stations", "Ranging"),
                ("Ground", "Mission Control", "Telemetry processing"), ("Ground", "Mission Control", "Commanding"),
                ("Ground", "Data Processing Centers", "Mission Analysis"), ("Ground", "Remote Terminals", "Network access"),
                ("Ground", "User Ground Segment", "Development"), ("Space", "Platform", "Bus"),
                ("Space", "Payload", "Payload Data Handling Systems"), ("Link", "Link", "Between Platform and Payload"),
                ("User", "User", "Transmission")
            ]

    def load_controls_from_csv(self):
        """Load controls from Control.csv"""
        controls_file = os.path.join(get_base_path(), "Control.csv")
        self.CONTROLS = []
        
        try:
            with open(controls_file, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                for row in reader:
                    control_id = row.get('#', '').strip()
                    control_cluster = row.get('Control cluster', '').strip()
                    control_title = row.get('Control title', '').strip()
                    control_description = row.get('Control', '').strip()
                    threats_addressed = row.get('Threats addressed', '').strip()
                    criteria = row.get('Criterio', '').strip()
                    segment = row.get('Segment', '').strip()
                    reference = row.get('Reference frameworks', '').strip()  # Add reference column
                    lifecycle = row.get('Lifecycle phase', '').strip()  # Add lifecycle column
                    
                    if control_id and control_title:
                        self.CONTROLS.append({
                            'id': control_id,
                            'cluster': control_cluster,
                            'title': control_title,
                            'description': control_description,
                            'threats_addressed': threats_addressed,
                            'criteria': criteria,
                            'segment': segment,
                            'reference': reference,  # Add reference field
                            'lifecycle': lifecycle   # Add lifecycle field
                        })
            
            print(f"[OK] Loaded {len(self.CONTROLS)} controls from {controls_file}")
            
        except FileNotFoundError:
            print(f"[ERROR] File not found: {controls_file}")
            self.CONTROLS = []
        except Exception as e:
            print(f"[ERROR] Error loading controls: {e}")
            self.CONTROLS = []

    def create_interface(self):
        """Creates the main interface"""
        # Header
        header = tk.Frame(self.root, bg=self.COLORS['light'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        tk.Label(header, text="Risk Assessment Tool for Advanced Phases", 
                font=('Segoe UI', 16, 'bold'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(pady=15)
        
        # Main container
        main_frame = tk.Frame(self.root, bg=self.COLORS['white'])
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Threats table
        self.create_threats_table(main_frame)
        
        # Buttons
        self.create_buttons(main_frame)

    def create_threats_table(self, parent):
        """Creates the threats table with scroll support"""
        # Main container with scroll
        main_container = tk.Frame(parent, bg=self.COLORS['white'])
        main_container.pack(fill='both', expand=True)

        # Canvas and scrollbar for main table
        canvas = tk.Canvas(main_container, bg=self.COLORS['white'])
        scrollbar = tk.Scrollbar(main_container, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.COLORS['white'])

        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Setup global mouse wheel scrolling for the main table
        self.setup_global_mousewheel(scrollable_frame, canvas)

        # Threat table
        table_frame = tk.LabelFrame(scrollable_frame, text="Threat Risk Assessment",
                                   font=('Segoe UI', 12, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=5, pady=5)
        table_frame.pack(fill='both', expand=True)

        # Headers
        headers = ["Threat", "Likelihood", "Impact", "Risk"]
        for j, header in enumerate(headers):
            cell = tk.Label(table_frame, text=header,
                           font=('Segoe UI', 11, 'bold'),
                           bg=self.COLORS['primary'], fg=self.COLORS['white'],
                           relief='ridge', bd=1, wraplength=400)
            cell.grid(row=0, column=j, sticky='ew', padx=1, pady=1, ipady=8)
        
        # Data Rows
        self.threat_cells = {}
        for i, threat in enumerate(self.THREATS, 1):
            # Threat name
            name_cell = tk.Label(table_frame, text=threat,
                               font=('Segoe UI', 10),
                               bg=self.COLORS['white'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, anchor='w',
                               wraplength=500)
            name_cell.grid(row=i, column=0, sticky='ew', padx=1, pady=1, ipady=5)
            
            # Initialize cells dictionary for this threat
            self.threat_cells[threat] = {}
            
            # Likelihood, Impact, Risk cells
            for j, cell_type in enumerate(['likelihood', 'impact', 'risk'], 1):
                cell = tk.Label(table_frame, text="",
                               font=('Segoe UI', 10),
                               bg=self.COLORS['white'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1)
                cell.grid(row=i, column=j, sticky='ew', padx=1, pady=1, ipady=5)
                self.threat_cells[threat][cell_type] = cell
        
        # Grid configuration
        table_frame.grid_columnconfigure(0, weight=2, minsize=400)  # Threat column
        table_frame.grid_columnconfigure(1, weight=1, minsize=220)  # Likelihood column
        table_frame.grid_columnconfigure(2, weight=1, minsize=220)  # Impact column
        table_frame.grid_columnconfigure(3, weight=1, minsize=220)  # Risk column
        
        # Ensure the table frame and all its cells also have mouse wheel scrolling
        self.setup_global_mousewheel(table_frame, canvas)
        
        # Specifically ensure all table cells have mouse wheel scrolling
        self.ensure_mousewheel_on_table_cells(canvas)

    def create_buttons(self, parent):
        """Creates the buttons"""
        button_frame = tk.Frame(parent, bg=self.COLORS['white'])
        button_frame.pack(fill='x', pady=(10, 0))

        # Container for buttons side by side
        buttons_container = tk.Frame(button_frame, bg=self.COLORS['white'])
        buttons_container.pack(pady=10)

        # IMPORT REPORT 0-A button (first)
        legacy_import_btn = tk.Button(buttons_container, text="IMPORT REPORT 0-A",
                                     font=('Segoe UI', 12, 'bold'),
                                     bg="#7e2929", fg=self.COLORS['white'],
                                     relief='flat', padx=30, pady=10,
                                     command=self.export_import_manager.legacy_import)
        legacy_import_btn.pack(side='left', padx=(0, 10))

        # IMPORT REPORT button (second)
        import_report_btn = tk.Button(buttons_container, text="IMPORT REPORT",
                                    font=('Segoe UI', 12, 'bold'),
                                    bg='#2c3e50', fg=self.COLORS['white'],
                                    relief='flat', padx=30, pady=10,
                                    command=self.export_import_manager.import_word_report)
        import_report_btn.pack(side='left', padx=(0, 10))

        # ASSET ANALYSIS button (third)
        add_asset_btn = tk.Button(buttons_container, text="ASSET ANALYSIS",
                                 font=('Segoe UI', 12, 'bold'),
                                 bg=self.COLORS['success'], fg=self.COLORS['white'],
                                 relief='flat', padx=30, pady=10,
                                 command=self.open_asset_window)
        add_asset_btn.pack(side='left', padx=(0, 10))

        # THREAT ANALYSIS button (fourth)
        self.threat_analysis_btn = tk.Button(buttons_container, text="THREAT ANALYSIS",
                                           font=('Segoe UI', 12, 'bold'),
                                           bg=self.COLORS['primary'], fg=self.COLORS['white'],
                                           relief='flat', padx=30, pady=10,
                                           command=self.open_threat_window)
        self.threat_analysis_btn.pack(side='left', padx=(0, 10))

        # CONTROLS MANAGEMENT button (fifth) - NEW!
        controls_btn = tk.Button(buttons_container, text="CONTROLS MANAGEMENT",
                                font=('Segoe UI', 12, 'bold'),
                                bg='#9b59b6', fg=self.COLORS['white'],
                                relief='flat', padx=30, pady=10,
                                command=self.open_controls_window)
        controls_btn.pack(side='left', padx=(0, 10))

        # EXPORT CSV button (sixth)
        export_btn = tk.Button(buttons_container, text="EXPORT CSV",
                              font=('Segoe UI', 12, 'bold'),
                              bg='#e67e22', fg=self.COLORS['white'],
                              relief='flat', padx=30, pady=10,
                              command=self.export_import_manager.export_csv)
        export_btn.pack(side='left', padx=(0, 10))

        # EXPORT REPORT button (seventh and last)
        export_report_btn = tk.Button(buttons_container, text="EXPORT REPORT",
                                     font=('Segoe UI', 12, 'bold'),
                                     bg='#8e44ad', fg=self.COLORS['white'],
                                     relief='flat', padx=30, pady=10,
                                     command=self.export_import_manager.export_word_report)
        export_report_btn.pack(side='left', padx=(0, 0))

    def update_threat_analysis_button_state(self):
        """Update the visual state of the Threat Analysis button based on active controls"""
        if hasattr(self, 'threat_analysis_btn') and hasattr(self, 'applied_controls'):
            if self.applied_controls:
                # Controls are active - make button orange to indicate read-only mode
                self.threat_analysis_btn.config(
                    bg='#f39c12',  # Orange color to indicate read-only
                    fg='white',    # White text
                    state='normal'
                )
            else:
                # No controls active - restore original color for full editing
                self.threat_analysis_btn.config(
                    bg=self.COLORS['primary'],  # Original blue color
                    fg=self.COLORS['white'],    # White text
                    state='normal'
                )

    def open_threat_window(self):
        """Open Threat Analysis window"""
        # Check if there are active controls - open in read-only mode if present
        read_only_mode = hasattr(self, 'applied_controls') and bool(self.applied_controls)
        
        if read_only_mode:
            # Show info message for read-only mode
            messagebox.showinfo(
                "Read-Only Mode",
                f"Threat Analysis opened in READ-ONLY mode.\n\n"
                f"Currently active controls: {len(self.applied_controls)}\n\n"
                f"You can view existing assessments but cannot modify criteria.\n"
                f"To enable editing, go to 'CONTROLS MANAGEMENT' and click 'CLEAR ALL CONTROLS'.",
                icon='info'
            )
        
        window = tk.Toplevel(self.root)
        window.title("Threat Analysis" + (" - READ ONLY" if read_only_mode else ""))
        window.geometry("1670x800")
        window.configure(bg=self.COLORS['white'])
        window.transient(self.root)
        window.grab_set()
        
        # Header with read-only indicator
        header = tk.Frame(window, bg=self.COLORS['light'] if not read_only_mode else '#f39c12', height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        header_text = "Threat Analysis"
        if read_only_mode:
            header_text += " - READ ONLY MODE"
        
        tk.Label(header, text=header_text,
                font=('Segoe UI', 14, 'bold'),
                bg=self.COLORS['light'] if not read_only_mode else '#f39c12', 
                fg=self.COLORS['dark'] if not read_only_mode else 'white').pack(pady=12)
        
        # Store read-only state for use in create_threat_content
        self.threat_window_read_only = read_only_mode
        
        # Main content with scroll
        self.create_threat_content(window)

    def open_asset_window(self):
        """Open Asset Analysis window"""
        window = tk.Toplevel(self.root)
        window.title("Asset Analysis")
        window.geometry("1600x800")
        window.configure(bg=self.COLORS['white'])
        window.transient(self.root)
        window.grab_set()
        
        # Header
        header = tk.Frame(window, bg=self.COLORS['light'], height=50)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        # Header content with title and help button
        header_content = tk.Frame(header, bg=self.COLORS['light'])
        header_content.pack(expand=True, fill='x', pady=12)
        
        tk.Label(header_content, text="Asset Analysis",
                font=('Segoe UI', 14, 'bold'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(side='left', padx=(20, 0))
        
        

        # Main content with scroll
        self.create_asset_content(window)

    def open_controls_window(self):
        """Open Controls Management window"""
        self.controls_window = tk.Toplevel(self.root)
        self.controls_window.title("🛡️ Controls Management - Dynamic Impact Analysis")
        self.controls_window.geometry("1680x950")  # Optimized for 1920x1080 screens
        self.controls_window.configure(bg=self.COLORS['white'])
        self.controls_window.transient(self.root)
        self.controls_window.grab_set()
        
        # Header (style consistent with assessment windows)
        header = tk.Frame(self.controls_window, bg=self.COLORS['primary'], height=60)
        header.pack(fill='x')
        header.pack_propagate(False)
        
        # Header content
        header_content = tk.Frame(header, bg=self.COLORS['primary'])
        header_content.pack(expand=True, fill='both', padx=20, pady=12)
        
        # Title
        tk.Label(header_content, text="🛡️ Controls Management & Dynamic Impact Analysis",
                font=('Segoe UI', 14, 'bold'),
                bg=self.COLORS['primary'], fg=self.COLORS['white']).pack(anchor='w')
        
        tk.Label(header_content, text="Select security controls to apply and see their dynamic impact on threat criteria",
                font=('Segoe UI', 10),
                bg=self.COLORS['primary'], fg=self.COLORS['light']).pack(anchor='w', pady=(2, 0))

        # Main content with scroll
        self.create_controls_content(self.controls_window)
        
        # Initialize interface state after creation
        self.controls_window.after(100, self.update_impact_analysis)

    def show_controls_help(self):
        """Show comprehensive help dialog for Controls Management"""
        help_window = tk.Toplevel(self.controls_window)
        help_window.title("Controls Management - Help")
        help_window.geometry("1200x700")
        help_window.configure(bg=self.COLORS['white'])
        help_window.resizable(True, True)
        
        # Center the window (consistent with other help windows)
        help_window.transient(self.controls_window)
        help_window.grab_set()
        
        # Title (consistent style with other help windows)
        title_label = tk.Label(help_window, text="Controls Management - User Guide", 
                              font=('Segoe UI', 14, 'bold'),
                              bg=self.COLORS['white'], fg=self.COLORS['dark'])
        title_label.pack(pady=(20, 15))
        
        # Create scrollable frame for the content (consistent structure)
        canvas = tk.Canvas(help_window, bg=self.COLORS['white'], highlightthickness=0)
        scrollbar = tk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Add separator and tool explanation section (consistent with other help windows)
        separator_frame = tk.Frame(scrollable_frame, bg=self.COLORS['gray'], height=2)
        separator_frame.pack(fill='x', pady=(0, 15), padx=15)
        
        # Tool explanation title
        explanation_title = tk.Label(scrollable_frame, text="How the Controls Management System Works", 
                                    font=('Segoe UI', 13, 'bold'),
                                    bg=self.COLORS['white'], fg=self.COLORS['primary'])
        explanation_title.pack(pady=(10, 15), padx=15, anchor='w')
        
        # Tool explanation content
        explanation_frame = tk.Frame(scrollable_frame, bg=self.COLORS['light'], relief='ridge', bd=1)
        explanation_frame.pack(fill='x', padx=15, pady=(0, 20))
        
        explanation_text = """The Controls Management system for Phase B-C-D allows you to select and apply security controls to your risk assessment. It provides intelligent asset compatibility filtering, real-time impact analysis, and integrates seamlessly with your threat analysis. Here's how to use it effectively:

1. ACCESSING CONTROLS MANAGEMENT:
   • From the main Risk Assessment window, click "CONTROLS MANAGEMENT" to open the controls selection interface
   • The system loads all available security controls with their compatibility information
   • Controls are organized by clusters (families) for easier navigation
   • Use this feature to enhance your risk assessment with implemented security measures

2. BROWSING AND SEARCHING CONTROLS:
   • Use the search bar at the top to filter controls by keywords, threat names, or criteria
   • Search examples: "jamming", "detection", "access control", "encryption"
   • Expand cluster categories by clicking the header to see all controls in that family
   • All controls show their ID, title, description, and compatibility information
   • Clear search to return to the full list of available controls

3. INTELLIGENT ASSET COMPATIBILITY:
   • Controls are automatically filtered based on asset segment compatibility
   • GROUND: Ground stations, mission control, data processing centers, terrestrial infrastructure
   • SPACE: Platform, payload, satellite systems, on-orbit assets
   • LINK: Communication links, data transmission channels, network connections
   • USER: User terminals, interfaces, end-user systems
   • HUMAN RESOURCES: Personnel controls, organizational measures, training
   • Only compatible controls will affect each asset during risk calculation

4. SELECTING CONTROLS FOR APPLICATION:
   • Check the boxes next to controls you want to apply to your assessment
   • Selected controls are immediately reflected in the real-time impact analysis panel
   • Use "Clear All Controls" button to deselect all currently selected controls
   • Controls remain selected until you apply them or clear the selection
   • MOUSE WHEEL PROTECTION: Scrolling won't accidentally change your selections

5. REAL-TIME IMPACT ANALYSIS:
   • The right panel provides live analysis of your control selections
   • CRITERIA IMPACT: Shows which threat assessment criteria will be improved
   • THREAT COVERAGE: Displays which threats are addressed by your selected controls
   • COVERAGE LEVELS: Excellent (4+ controls), Good (2-3 controls), Basic (1 control)
   • CONTROL SUMMARY: Organized breakdown by control clusters and effectiveness
   • ASSET IMPACT: Shows how controls will affect each asset segment

6. APPLYING CONTROLS TO YOUR ASSESSMENT:
   • Click "SAVE & APPLY CONTROLS" to integrate selected controls with your threat analysis
   • Selected controls are permanently saved to your assessment configuration
   • Threat criteria scores are automatically updated based on control effectiveness
   • Risk calculations are recalculated with control impact factors
   • Main threat assessment table reflects new risk levels with controls applied

7. INTEGRATION WITH THREAT ANALYSIS:
   • Once controls are applied, the Threat Analysis window enters READ-ONLY MODE
   • READ-ONLY MODE: Threat assessment scores become view-only to preserve control integration
   • This prevents accidental modification of assessments that have been enhanced with controls
   • To modify threat assessments after applying controls, you must clear controls first
   • This ensures consistency between your control selections and threat evaluations

8. CONTROL MANAGEMENT WORKFLOW:
   • STEP 1: Complete your initial threat assessment without controls
   • STEP 2: Use Controls Management to select appropriate security measures
   • STEP 3: Review impact analysis to ensure adequate threat coverage
   • STEP 4: Apply controls to see updated risk levels with security measures
   • STEP 5: Generate final reports with control-enhanced risk assessments

9. ADVANCED FEATURES:
   • DYNAMIC CONTENT SCROLLING: Enhanced mouse wheel handling for smooth navigation
   • PERSISTENT SELECTIONS: Your control choices are maintained during session
   • COMPATIBILITY VALIDATION: System prevents invalid control applications
   • IMPACT PREVIEW: See the effects before committing to control application
   • CLUSTER ORGANIZATION: Controls grouped by security domains for logical selection

10. BEST PRACTICES:
   • Complete your baseline threat assessment before applying controls
   • Select controls that address your highest-risk threats first
   • Ensure good coverage across all critical asset segments
   • Review the impact analysis to identify coverage gaps
   • Consider the cumulative effect of multiple controls on the same threat
   • Apply controls incrementally and review their impact on overall risk posture
   • Generate reports both before and after control application for comparison

11. IMPORTANT NOTES:
   • COMPATIBILITY: Controls only apply to compatible asset segments automatically
   • READ-ONLY PROTECTION: Threat Analysis becomes read-only after applying controls
   • CLEAR CONTROLS: Use "Clear All Controls" to reset and allow threat assessment modifications
   • REAL-TIME UPDATES: Impact analysis updates immediately as you select/deselect controls
   • PERSISTENT STATE: Control selections are maintained until explicitly applied or cleared"""
        
        explanation_label = tk.Label(explanation_frame, text=explanation_text,
                                   font=('Segoe UI', 10),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   anchor='nw', justify='left', wraplength=1150,
                                   padx=20, pady=15)
        explanation_label.pack(fill='both', expand=True)
        
        # Pack canvas and scrollbar (consistent with other help windows)
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        # Setup global mouse wheel scrolling for the help window (consistent)
        self.setup_global_mousewheel(scrollable_frame, canvas)
        
        # Focus on help window
        help_window.focus_set()

    def create_controls_content(self, window):
        """Creates the controls management content window"""
        # Initialize control selection storage
        self.selected_controls = {}
        self.control_impact_display = {}
        
        # Main container
        main_container = tk.Frame(window, bg=self.COLORS['white'])
        main_container.pack(fill='both', expand=True, padx=15, pady=10)
        
        # Create horizontal layout with fixed proportions using grid
        content_frame = tk.Frame(main_container, bg=self.COLORS['white'])
        content_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        # Configure grid columns with weights: 60% vs 40% (optimized for 1920x1080)
        content_frame.grid_columnconfigure(0, weight=12, minsize=820)   # Left panel - 60% weight, reduced width
        content_frame.grid_columnconfigure(1, weight=8, minsize=420)    # Right panel - 40% weight, reduced width
        content_frame.grid_rowconfigure(0, weight=1)
        
        # Left panel: Controls selection (60% width)
        left_panel = tk.Frame(content_frame, bg=self.COLORS['white'])
        left_panel.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        
        # Right panel: Impact analysis (40% width)  
        right_panel = tk.Frame(content_frame, bg=self.COLORS['light'])
        right_panel.grid(row=0, column=1, sticky='nsew')
        
        # === LEFT PANEL: CONTROLS SELECTION ===
        self.create_controls_selection_panel(left_panel)

        # === RIGHT PANEL: IMPACT ANALYSIS ===
        self.create_controls_impact_panel(right_panel)
        
        # === BOTTOM: BUTTON BAR - Always visible ===
        bottom_button_frame = tk.Frame(main_container, bg=self.COLORS['light'], relief='ridge', bd=2)
        bottom_button_frame.pack(fill='x', pady=(15, 0))
        
        # Create button container with consistent styling (like assessment windows)
        button_container = tk.Frame(bottom_button_frame, bg=self.COLORS['light'])
        button_container.pack(expand=True, pady=15)
        
        # Clear All Controls button (left side)
        clear_all_btn = tk.Button(button_container, text="🗑️ Clear All Controls",
                                 font=('Segoe UI', 11, 'bold'),
                                 bg='#e74c3c', fg=self.COLORS['white'],
                                 relief='flat', padx=20, pady=10,
                                 command=self.clear_all_controls_gui)
        clear_all_btn.pack(side='left', padx=(0, 15))
        
        # Help button (left side, next to clear)
        help_btn = tk.Button(button_container, text="❓ Help",
                            font=('Segoe UI', 11, 'bold'),
                            bg='#8e44ad', fg=self.COLORS['white'],
                            relief='flat', padx=20, pady=10,
                            command=self.show_controls_help)
        help_btn.pack(side='left', padx=(0, 30))
        
        # Save & Apply button (right side)
        save_apply_btn = tk.Button(button_container, text="💾 Save & Apply Controls",
                                  font=('Segoe UI', 12, 'bold'),
                                  bg='#27ae60', fg=self.COLORS['white'],
                                  relief='flat', padx=30, pady=10,
                                  command=self.save_and_apply_controls)
        save_apply_btn.pack(side='right', padx=(0, 0))

    def create_controls_selection_panel(self, parent):
        """Creates the controls selection panel"""
        # Header
        header_frame = tk.Frame(parent, bg=self.COLORS['primary'], height=45)
        header_frame.pack(fill='x', pady=(0, 12))
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📋 Available Security Controls",
                font=('Segoe UI', 12, 'bold'),
                bg=self.COLORS['primary'], fg=self.COLORS['white']).pack(pady=12)
        
        # Search frame (style consistent with assessment windows)
        search_frame = tk.Frame(parent, bg=self.COLORS['light'], relief='ridge', bd=1)
        search_frame.pack(fill='x', pady=(0, 10))
        
        # Search label and entry
        search_label_frame = tk.Frame(search_frame, bg=self.COLORS['light'])
        search_label_frame.pack(fill='x', padx=15, pady=10)
        
        tk.Label(search_label_frame, text="🔍 Search Controls:",
                font=('Segoe UI', 10, 'bold'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(side='left')
        
        # Create search entry with consistent styling
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(search_label_frame, textvariable=self.search_var,
                                    font=('Segoe UI', 10),
                                    bg=self.COLORS['white'], fg=self.COLORS['dark'],
                                    relief='solid', bd=1)
        self.search_entry.pack(side='left', fill='x', expand=True, padx=(15, 0))
        
        # Bind search functionality
        self.search_var.trace('w', self.on_search_changed)
        
        # Search instructions
        search_instruction_frame = tk.Frame(search_frame, bg=self.COLORS['light'])
        search_instruction_frame.pack(fill='x', padx=15, pady=(0, 8))
        
        tk.Label(search_instruction_frame, text="💡 Type to search by threat name, control title, or description. Clear to show all controls.",
                font=('Segoe UI', 9, 'italic'),
                bg=self.COLORS['light'], fg=self.COLORS['gray']).pack()
        
        # Instructions (updated)
        instruction_frame = tk.Frame(parent, bg=self.COLORS['light'], relief='ridge', bd=1)
        instruction_frame.pack(fill='x', pady=(0, 15))
        
        tk.Label(instruction_frame, text="💡 Click on cluster names to expand and see individual controls. Use search to filter controls.",
                font=('Segoe UI', 10, 'italic'),
                bg=self.COLORS['light'], fg=self.COLORS['dark']).pack(pady=8, padx=15)
        
        # Scrollable controls list
        canvas = tk.Canvas(parent, bg=self.COLORS['white'], highlightthickness=0)
        scrollbar = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        
        # Store references for search functionality
        self.controls_canvas = canvas
        self.controls_scrollable_frame = scrollable_frame
        
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        # Store all controls data for search
        self.all_controls_data = self._build_controls_data()
        
        # Create initial display (all controls)
        self.display_controls(self.all_controls_data)
        
        # Setup mouse wheel scrolling - ONLY for the controls list
        self.setup_controls_mousewheel(scrollable_frame, canvas)
        
        # Remove the global window binding to prevent conflicts
        # We'll handle scroll events specifically per section
    
    def _build_controls_data(self):
        """Build controls data structure for display and search"""
        # Group controls by cluster
        controls_by_cluster = {}
        for control in self.CONTROLS:
            cluster = control['cluster']
            if cluster not in controls_by_cluster:
                controls_by_cluster[cluster] = []
            controls_by_cluster[cluster].append(control)
        
        return controls_by_cluster
    
    def display_controls(self, controls_by_cluster):
        """Display controls in the scrollable frame"""
        # Clear existing content
        for widget in self.controls_scrollable_frame.winfo_children():
            widget.destroy()
        
        # Create expandable sections for each cluster with better styling
        for cluster_name, cluster_controls in controls_by_cluster.items():
            if not cluster_controls:  # Skip empty clusters
                continue
                
            # Cluster header with gradient-like effect - full width
            cluster_frame = tk.Frame(self.controls_scrollable_frame, bg='#6c7ae0', relief='flat', bd=0)
            cluster_frame.pack(fill='x', pady=(8, 0), padx=0)  # No padding at all to maximize width
            
            # Use a Button instead of Checkbutton for better control over width
            cluster_var = tk.BooleanVar(value=False)
            
            # Create the button first - directly in scrollable frame for full width
            cluster_btn = tk.Button(self.controls_scrollable_frame,
                                   text=f"📁 {cluster_name} ({len(cluster_controls)} controls)",
                                   font=('Segoe UI', 11, 'bold'),
                                   bg='#6c7ae0', fg=self.COLORS['white'],
                                   activebackground='#5a67d8', activeforeground=self.COLORS['white'],
                                   relief='flat', bd=0,
                                   anchor='w')  # Left align text
            cluster_btn.pack(fill='x', expand=False, padx=0, pady=(8, 0))  # Fill entire width, no padding
            
            # Create a custom toggle function that references the button
            def create_toggle_function(cluster=cluster_name, var=cluster_var, btn=cluster_btn, controls_count=len(cluster_controls)):
                def toggle_cluster():
                    current = var.get()
                    var.set(not current)
                    self.toggle_cluster_controls(cluster, not current)
                    # Update button text to show state
                    btn_text = f"📂 {cluster} ({controls_count} controls)" if not current else f"📁 {cluster} ({controls_count} controls)"
                    btn.config(text=btn_text)
                return toggle_cluster
            
            # Assign the command
            cluster_btn.config(command=create_toggle_function())
            
            # Controls container (initially hidden)
            controls_container = tk.Frame(self.controls_scrollable_frame, bg='#f8f9fa', relief='flat', bd=1)
            controls_container.pack(fill='x', padx=15)  # Reduced padding from 20 to 15
            
            # Store cluster info for toggling
            setattr(self, f"cluster_{cluster_name.replace(' ', '_')}_container", controls_container)
            setattr(self, f"cluster_{cluster_name.replace(' ', '_')}_var", cluster_var)
            
            # Initially hide
            controls_container.pack_forget()
            
            # Add individual controls
            for control in cluster_controls:
                self.create_control_item(controls_container, control)
    
    def on_search_changed(self, *args):
        """Handle search text changes"""
        search_text = self.search_var.get().lower().strip()
        
        if not search_text:
            # Show all controls
            self.display_controls(self.all_controls_data)
        else:
            # Filter controls based on search text
            filtered_controls = {}
            
            for cluster_name, cluster_controls in self.all_controls_data.items():
                matching_controls = []
                
                for control in cluster_controls:
                    # Search in multiple fields
                    title = control.get('title', '').lower()
                    description = control.get('description', '').lower()
                    threats = control.get('threats_addressed', '').lower()
                    criteria = control.get('criteria', '').lower()
                    
                    if (search_text in title or 
                        search_text in description or 
                        search_text in threats or 
                        search_text in criteria or
                        search_text in cluster_name.lower()):
                        matching_controls.append(control)
                
                if matching_controls:
                    filtered_controls[cluster_name] = matching_controls
            
            self.display_controls(filtered_controls)
        
    def create_control_item(self, parent, control):
        """Creates a single control item with checkbox and info"""
        # Check if control is already applied
        is_applied = self.is_control_applied(control['id'])
        
        # Set background color based on status
        bg_color = '#e8f5e8' if is_applied else self.COLORS['white']
        
        control_frame = tk.Frame(parent, bg=bg_color, relief='ridge', bd=1)
        control_frame.pack(fill='x', pady=2, padx=5)
        
        # Control selection checkbox
        control_var = tk.BooleanVar()
        control_var.set(is_applied)  # Pre-select if already applied
        self.selected_controls[control['id']] = control_var
        
        # Main control info
        info_frame = tk.Frame(control_frame, bg=bg_color)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        # Checkbox and title
        check_frame = tk.Frame(info_frame, bg=bg_color)
        check_frame.pack(fill='x')
        
        # Add status indicator to title
        title_text = f"#{control['id']} - {control['title']}"
        if is_applied:
            title_text += "  APPLIED"
        
        checkbox = tk.Checkbutton(check_frame,
                                 text=title_text,
                                 font=('Segoe UI', 10, 'bold'),
                                 bg=bg_color, 
                                 fg='#27ae60' if is_applied else self.COLORS['dark'],
                                 variable=control_var,
                                 command=lambda c=control: self.on_control_selected(c))
        checkbox.pack(anchor='w')
        
        # Disable mouse wheel on checkbox to prevent accidental changes
        self.disable_mousewheel_on_checkbox(checkbox)
        
        # Description (truncated)
        desc_text = control['description'][:400] + "..." if len(control['description']) > 400 else control['description']
        desc_label = tk.Label(info_frame, text=desc_text,
                             font=('Segoe UI', 9),
                             bg=bg_color, fg='#666666',
                             wraplength=1000, justify='left')
        desc_label.pack(anchor='w', pady=(2, 0))
        
        # Threats and criteria info
        if control['threats_addressed'] or control['criteria'] or control.get('segment'):
            impact_frame = tk.Frame(info_frame, bg='#f8f9fa', relief='flat', bd=1)
            impact_frame.pack(fill='x', pady=(5, 0))
            
            # Show segment information first if available
            if control.get('segment'):
                segment_text = f"🏢 Segment: {control['segment']}"
                # Special display for Human -> User Ground Segment
                if control['segment'].lower() == 'human':
                    segment_text += " (User Ground Segment)"
                
                segment_label = tk.Label(impact_frame, 
                                       text=segment_text,
                                       font=('Segoe UI', 9, 'bold'),
                                       bg='#f8f9fa', fg='#6c5ce7',
                                       wraplength=1000, justify='left')
                segment_label.pack(anchor='w', padx=5, pady=2)
            
            if control['threats_addressed']:
                threats_label = tk.Label(impact_frame, 
                                       text=f" Threats: {control['threats_addressed'][:200]}...",
                                       font=('Segoe UI', 9, 'italic'),
                                       bg='#f8f9fa', fg='#d63031',
                                       wraplength=1000, justify='left')
                threats_label.pack(anchor='w', padx=5, pady=2)
            
            if control['criteria']:
                criteria_label = tk.Label(impact_frame, 
                                        text=f" Criteria Impact: {control['criteria']}",
                                        font=('Segoe UI', 9, 'italic'),
                                        bg='#f8f9fa', fg='#00b894',
                                        wraplength=1000, justify='left')  # Reduced from 1200 to 1000
                criteria_label.pack(anchor='w', padx=5, pady=2)
    
    def toggle_cluster_controls(self, cluster_name, is_expanded):
        """Toggle visibility of controls in a cluster"""
        container_attr = f"cluster_{cluster_name.replace(' ', '_')}_container"
        controls_container = getattr(self, container_attr, None)
        
        if controls_container:
            if is_expanded:
                controls_container.pack(fill='x', padx=15)
            else:
                controls_container.pack_forget()
    
    def create_controls_impact_panel(self, parent):
        """Creates the impact analysis panel"""
        # Header (consistent with assessment windows)
        header_frame = tk.Frame(parent, bg=self.COLORS['primary'], height=45)
        header_frame.pack(fill='x', pady=(0, 10))
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="⚡ Dynamic Impact Analysis",
                font=('Segoe UI', 12, 'bold'),
                bg=self.COLORS['primary'], fg=self.COLORS['white']).pack(pady=11)
        
        # Stats frame
        stats_frame = tk.Frame(parent, bg=self.COLORS['white'], relief='ridge', bd=1)
        stats_frame.pack(fill='x', pady=(0, 10), padx=5)
        
        self.stats_label = tk.Label(stats_frame, text=" No controls selected yet",
                                   font=('Segoe UI', 10, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['gray'],
                                   pady=10)
        self.stats_label.pack()
        
        # Impact visualization canvas
        canvas = tk.Canvas(parent, bg=self.COLORS['white'], highlightthickness=0)
        scrollbar_impact = tk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        impact_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        
        impact_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=impact_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar_impact.set)
        
        canvas.pack(side="left", fill="both", expand=True, padx=(5, 0))
        scrollbar_impact.pack(side="right", fill="y")
        
        self.impact_display_frame = impact_frame
        
        # Initial empty state
        self.show_empty_impact_state()
        
        # Setup mouse wheel scrolling - specific for impact panel
        self.setup_impact_mousewheel(impact_frame, canvas)
    
    def show_empty_impact_state(self):
        """Shows empty state for impact analysis"""
        # Clear existing content
        for widget in self.impact_display_frame.winfo_children():
            widget.destroy()
        
        empty_frame = tk.Frame(self.impact_display_frame, bg=self.COLORS['white'])
        empty_frame.pack(fill='both', expand=True, pady=50)
        
        tk.Label(empty_frame, text="🔍",
                font=('Segoe UI', 48),
                bg=self.COLORS['white'], fg=self.COLORS['gray']).pack()
        
        tk.Label(empty_frame, text="Select security controls to see their impact",
                font=('Segoe UI', 12),
                bg=self.COLORS['white'], fg=self.COLORS['gray']).pack(pady=(10, 0))
        
        tk.Label(empty_frame, text="Controls will dynamically affect threat criteria scores",
                font=('Segoe UI', 10),
                bg=self.COLORS['white'], fg='#95a5a6').pack(pady=(5, 0))
    
    def on_control_selected(self, control):
        """Handle control selection/deselection"""
        self.update_impact_analysis()
    
    def update_impact_analysis(self):
        """Updates the impact analysis display"""
        # Get selected controls
        selected = []
        for control_id, var in self.selected_controls.items():
            if var.get():
                # Find control data
                control_data = next((c for c in self.CONTROLS if c['id'] == control_id), None)
                if control_data:
                    selected.append(control_data)
        
        # Update stats
        if not selected:
            self.stats_label.config(text=" No controls selected yet")
            self.show_empty_impact_state()
            return
        
        self.stats_label.config(text=f" {len(selected)} controls selected • Analyzing impact...")
        
        # Analyze impact
        self.analyze_and_display_impact(selected)
    
    def analyze_and_display_impact(self, selected_controls):
        """Analyzes and displays the impact of selected controls"""
        # Clear existing impact display
        for widget in self.impact_display_frame.winfo_children():
            widget.destroy()
        
        # Collect all impacted criteria and threats
        criteria_impact = {}
        threat_impact = {}
        
        for control in selected_controls:
            # Process criteria impact
            if control['criteria']:
                # Split by comma and clean up each criterion
                criteria_list = []
                for criterion in control['criteria'].split(','):
                    clean_criterion = criterion.strip()
                    if clean_criterion:  # Only add non-empty criteria
                        criteria_list.append(clean_criterion)
                
                for criterion in criteria_list:
                    if criterion not in criteria_impact:
                        criteria_impact[criterion] = []
                    criteria_impact[criterion].append(control)
            
            # Process threat impact
            if control['threats_addressed']:
                # Split threats by comma and clean up
                threats_list = []
                for threat in control['threats_addressed'].split(','):
                    clean_threat = threat.strip()
                    if clean_threat:
                        threats_list.append(clean_threat)
                        
                for threat in threats_list:
                    if threat not in threat_impact:
                        threat_impact[threat] = []
                    threat_impact[threat].append(control)
        
        # Display criteria impact
        if criteria_impact:
            criteria_section = self.create_criteria_impact_section(criteria_impact)
            if criteria_section:
                self.apply_impact_mousewheel_to_new_content(criteria_section)
        
        # Display threat impact  
        if threat_impact:
            threat_section = self.create_threat_impact_section(threat_impact)
            if threat_section:
                self.apply_impact_mousewheel_to_new_content(threat_section)
            
        # Display control summary
        control_summary = self.create_control_summary_section(selected_controls)
        if control_summary:
            self.apply_impact_mousewheel_to_new_content(control_summary)
    
    def create_criteria_impact_section(self, criteria_impact):
        """Creates the criteria impact visualization section"""
        section_frame = tk.LabelFrame(self.impact_display_frame, 
                                    text=" Criteria Impact Analysis",
                                    font=('Segoe UI', 12, 'bold'),
                                    bg=self.COLORS['white'], fg='#00b894',
                                    padx=10, pady=10)
        section_frame.pack(fill='x', pady=(0, 15), padx=5)
        
        # Threat criteria mapping
        threat_criteria_names = [
            "Vulnerability Effectiveness", "Mitigation Presence", "Detection Probability", 
            "Access Complexity", "Privilege Requirement", "Response Delay", "Resilience Impact"
        ]
        
        for criterion, controls in criteria_impact.items():
            # Create criterion row
            criterion_frame = tk.Frame(section_frame, bg='#e8f4fd', relief='ridge', bd=1)
            criterion_frame.pack(fill='x', pady=3)
            
            # Criterion name and impact
            header_frame = tk.Frame(criterion_frame, bg='#3498db')
            header_frame.pack(fill='x')
            
            # Check if this is a recognized threat criterion
            impact_level = len(controls)
            if criterion in threat_criteria_names:
                impact_color = '#27ae60' if impact_level >= 3 else '#f39c12' if impact_level >= 2 else '#3498db'
                impact_text = f"HIGH IMPACT" if impact_level >= 3 else f"MEDIUM IMPACT" if impact_level >= 2 else f"LOW IMPACT"
            else:
                impact_color = '#95a5a6'
                impact_text = f"GENERAL IMPACT"
            
            tk.Label(header_frame, text=f" {criterion}",
                    font=('Segoe UI', 11, 'bold'),
                    bg='#3498db', fg=self.COLORS['white']).pack(side='left', padx=10, pady=5)
            
            tk.Label(header_frame, text=f"{impact_text} ({len(controls)} controls)",
                    font=('Segoe UI', 10, 'bold'),
                    bg=impact_color, fg=self.COLORS['white']).pack(side='right', padx=10, pady=5)
            
            # Controls affecting this criterion
            controls_text = " • ".join([f"#{c['id']}: {c['title'][:30]}..." for c in controls[:3]])
            if len(controls) > 3:
                controls_text += f" • ... and {len(controls)-3} more"
            
            tk.Label(criterion_frame, text=controls_text,
                    font=('Segoe UI', 9),
                    bg='#e8f4fd', fg='#2c3e50',
                    wraplength=800, justify='left').pack(anchor='w', padx=10, pady=5)  # Increased from 700 to 800
        
        return section_frame
    
    def create_threat_impact_section(self, threat_impact):
        """Creates the threat impact visualization section"""
        section_frame = tk.LabelFrame(self.impact_display_frame, 
                                    text=" Threat Coverage Analysis",
                                    font=('Segoe UI', 12, 'bold'),
                                    bg=self.COLORS['white'], fg='#e74c3c',
                                    padx=10, pady=10)
        section_frame.pack(fill='x', pady=(0, 15), padx=5)
        
        # Sort threats by number of controls (most coverage first)
        sorted_threats = sorted(threat_impact.items(), key=lambda x: len(x[1]), reverse=True)
        
        for threat, controls in sorted_threats[:10]:  # Show top 10 most covered threats
            threat_frame = tk.Frame(section_frame, bg='#fdf2f2', relief='ridge', bd=1)
            threat_frame.pack(fill='x', pady=2)
            
            # Threat header
            header_frame = tk.Frame(threat_frame, bg='#e74c3c')
            header_frame.pack(fill='x')
            
            coverage_level = len(controls)
            coverage_color = '#27ae60' if coverage_level >= 4 else '#f39c12' if coverage_level >= 2 else '#e74c3c'
            coverage_text = f"EXCELLENT" if coverage_level >= 4 else f"GOOD" if coverage_level >= 2 else f"BASIC"
            
            tk.Label(header_frame, text=f" {threat[:50]}{'...' if len(threat) > 50 else ''}",
                    font=('Segoe UI', 10, 'bold'),
                    bg='#e74c3c', fg=self.COLORS['white']).pack(side='left', padx=10, pady=3)
            
            tk.Label(header_frame, text=f"{coverage_text} ({len(controls)} controls)",
                    font=('Segoe UI', 9, 'bold'),
                    bg=coverage_color, fg=self.COLORS['white']).pack(side='right', padx=10, pady=3)
            
            # Controls addressing this threat
            controls_text = " • ".join([f"#{c['id']}" for c in controls])
            
            tk.Label(threat_frame, text=f"Controls: {controls_text}",
                    font=('Segoe UI', 9),
                    bg='#fdf2f2', fg='#2c3e50').pack(anchor='w', padx=10, pady=3)
        
        return section_frame
    
    def create_control_summary_section(self, selected_controls):
        """Creates a summary section of selected controls"""
        section_frame = tk.LabelFrame(self.impact_display_frame, 
                                    text="📋 Selected Controls Summary",
                                    font=('Segoe UI', 12, 'bold'),
                                    bg=self.COLORS['white'], fg='#9b59b6',
                                    padx=10, pady=10)
        section_frame.pack(fill='x', pady=(0, 15), padx=5)
        
        # Group by cluster
        by_cluster = {}
        for control in selected_controls:
            cluster = control['cluster']
            if cluster not in by_cluster:
                by_cluster[cluster] = []
            by_cluster[cluster].append(control)
        
        for cluster, controls in by_cluster.items():
            cluster_frame = tk.Frame(section_frame, bg='#f8f9fa', relief='flat', bd=1)
            cluster_frame.pack(fill='x', pady=2)
            
            tk.Label(cluster_frame, text=f"📁 {cluster} ({len(controls)} controls)",
                    font=('Segoe UI', 10, 'bold'),
                    bg='#6c7ae0', fg=self.COLORS['white']).pack(anchor='w')
            
            for control in controls:
                tk.Label(cluster_frame, text=f"   • #{control['id']}: {control['title']}",
                        font=('Segoe UI', 9),
                        bg='#f8f9fa', fg='#2c3e50').pack(anchor='w', padx=10)
        
        return section_frame
    
    def save_and_apply_controls(self):
        """Apply/remove controls based on current selections"""
        # Check if we have threat data
        if not self.threat_data:
            messagebox.showwarning("No Data", "Please complete threat assessments first!")
            return
        
        # Get current UI selections and compare with applied controls
        newly_selected = []
        to_remove = []
        
        for control_id, var in self.selected_controls.items():
            is_selected = var.get()
            is_applied = self.is_control_applied(control_id)
            
            if is_selected and not is_applied:
                newly_selected.append(control_id)
            elif not is_selected and is_applied:
                to_remove.append(control_id)
        
        # Check if there are changes to apply
        if not newly_selected and not to_remove:
            return
        
        # Apply new controls
        applied_count = 0
        for control_id in newly_selected:
            control_data = self.get_control_by_id(control_id)
            if control_data and self.apply_control(control_id, control_data):
                applied_count += 1
        
        # Remove deselected controls
        removed_count = 0
        for control_id in to_remove:
            if self.remove_control(control_id):
                removed_count += 1
        
        # Update main table
        self.update_all_threats_in_main_table()
        
        # Show results
        total_active = len(self.applied_controls)
        if applied_count > 0 or removed_count > 0:            
            # Show summary dialog
            messagebox.showinfo("Controls Updated", 
                f"Control changes applied successfully!\n\n"
                f"• Applied: {applied_count} new controls\n"
                f"• Removed: {removed_count} controls\n"
                f"• Total active: {total_active} controls")
            
            # Update button state
            self.update_threat_analysis_button_state()
            
            # Auto-close after success
            self.controls_window.after(100, self.controls_window.destroy)
        else:
            pass  # No changes to apply
    
    def apply_controls_to_threats(self, selected_controls):
        """Apply controls impact to existing threat assessments"""
        impacts_applied = 0
        
        # Check if we have threat data to modify
        if not self.threat_data:
            messagebox.showinfo("No Data", "No threat assessment data found. Please complete threat analysis first.")
            return 0
        
        # Mapping of criteria names to threat criteria indices
        # Include multiple variations to handle inconsistent naming
        criteria_mapping = {
            "Vulnerability Effectiveness": 0,
            "Vulnerability effectiveness": 0,
            "vulnerability effectiveness": 0,
            "Mitigation Presence": 1,
            "mitigation presence": 1,
            "Detection Probability": 2,
            "detection probability": 2,
            "Access Complexity": 3,
            "access complexity": 3,
            "Privilege Requirement": 4,
            "privilege requirement": 4,
            "Response Delay": 5,
            "response delay": 5,
            "Resilience Impact": 6,
            "resilience impact": 6
        }
        
        print(f"[DEBUG] Starting to apply {len(selected_controls)} controls to threat data...")
        print(f"[DEBUG] Available threats in data: {list(self.threat_data.keys())}")
        
        for control in selected_controls:
            print(f" Processing control #{control['id']}: {control['title']}")
            
            # Process criteria impact
            if control['criteria']:
                # Split by comma and clean up each criterion
                criteria_list = []
                for criterion in control['criteria'].split(','):
                    # Strip all whitespace and normalize
                    clean_criterion = criterion.strip()
                    if clean_criterion:  # Only add non-empty criteria
                        criteria_list.append(clean_criterion)
                
                for criterion in criteria_list:
                    # Try exact match first, then try title case normalization
                    criterion_index = criteria_mapping.get(criterion)
                    if criterion_index is None:
                        # Try title case version
                        criterion_index = criteria_mapping.get(criterion.title())
                    if criterion_index is None:
                        # Try lowercase version
                        criterion_index = criteria_mapping.get(criterion.lower())
                    
                    if criterion_index is None:
                        print(f"   Unknown criterion: '{criterion}' - skipping")
                        continue  # Skip unknown criteria
                    
                    print(f"   Applying criterion '{criterion}' (index {criterion_index})")
                    
                    # Process threats addressed by this control
                    if control['threats_addressed']:
                        # Split threats by comma and clean up
                        threats_list = []
                        for threat in control['threats_addressed'].split(','):
                            clean_threat = threat.strip()
                            if clean_threat:
                                threats_list.append(clean_threat)
                        
                        for threat_name in threats_list:
                            # Find matching threat in our threat data
                            matching_threats = []
                            for threat in self.THREATS:
                                # More flexible matching - check if threat names contain each other
                                if (threat_name.lower() in threat.lower() or 
                                    threat.lower() in threat_name.lower() or
                                    self.normalize_threat_name(threat_name) == self.normalize_threat_name(threat)):
                                    matching_threats.append(threat)
                            
                            for threat in matching_threats:
                                if threat in self.threat_data:
                                    # Apply control impact (reduce criterion score)
                                    asset_data = self.threat_data[threat]
                                    
                                    for asset_key, scores in asset_data.items():
                                        # Check if this asset belongs to a segment that the control can be applied to
                                        is_compatible = self.is_asset_compatible_with_control(asset_key, control)
                                        
                                        if is_compatible:
                                            if isinstance(scores, dict) and str(criterion_index) in scores:
                                                # Reduce score by 1 (minimum 1, maximum 5)
                                                current_score = int(scores[str(criterion_index)])
                                                new_score = max(1, current_score - 1)
                                                scores[str(criterion_index)] = new_score
                                                impacts_applied += 1
                                                print(f"     Impact applied to {asset_key}: criterion {criterion_index} reduced from {current_score} to {new_score}")
        
        print(f" Total impacts applied: {impacts_applied}")
        return impacts_applied
    
    def is_asset_compatible_with_control(self, asset_key, control):
        """Check if an asset is compatible with a control based on segments"""
        # Get control segments (can be multiple, separated by comma)
        control_segments = []
        if control.get('segment'):
            # Split by comma and clean up each segment
            for segment in control['segment'].split(','):
                clean_segment = segment.strip()
                if clean_segment:
                    control_segments.append(clean_segment.lower())
        
        # If no segments specified in control, apply to all assets
        if not control_segments:
            return True
        
        # Get asset category from asset_key
        # asset_key should be in format "category_subcategory_asset" or similar
        asset_category = self.get_asset_category_from_key(asset_key)
        
        if not asset_category:
            return False  # Cannot determine asset category
        
        # Check compatibility based on segments
        asset_category_lower = asset_category.lower()
        
        print(f"[COMPAT] Asset '{asset_key}' (category: {asset_category_lower}) vs Control #{control['id']} (segments: {control_segments})")
        
        # Special mapping for Human Resources to Ground -> User Ground Segment
        if 'human resources' in control_segments:
            # For Human Resources controls, check if asset is in Ground -> User Ground Segment
            is_user_ground = self.is_user_ground_segment_asset(asset_key)
            if is_user_ground:
                print(f"[COMPAT]  Compatible via Human Resources -> User Ground Segment mapping")
                return True
        
        # Direct segment matching
        for control_segment in control_segments:
            if control_segment == asset_category_lower:
                print(f"[COMPAT]  Compatible: Direct match ({control_segment})")
                return True
            # Handle "human resources" mapping to ground
            if control_segment == 'human resources' and asset_category_lower == 'ground':
                is_user_ground = self.is_user_ground_segment_asset(asset_key)
                if is_user_ground:
                    print(f"[COMPAT]  Compatible via Human Resources -> Ground mapping")
                    return True
            # Handle "link" mapping (Link assets)
            if control_segment == 'link' and asset_category_lower == 'link':
                print(f"[COMPAT]  Compatible: Link segment match")
                return True
        
        print(f"[COMPAT]  Not compatible")
        return False
    
    def get_asset_category_from_key(self, asset_key):
        """Extract category from asset key"""
        # Handle numerical asset keys from threat analysis (e.g., "1_probability", "2_probability")
        if '_probability' in asset_key:
            try:
                index = int(asset_key.split('_')[0]) - 1  # Convert to 0-based index
                if 0 <= index < len(self.ASSET_CATEGORIES):
                    category, subcategory, asset = self.ASSET_CATEGORIES[index]
                    return category
            except (ValueError, IndexError):
                pass
        
        # Look for the asset in our ASSET_CATEGORIES list to get its category
        for category_tuple in self.ASSET_CATEGORIES:
            if len(category_tuple) >= 3:
                category, subcategory, asset = category_tuple
                # Build a potential asset key (this depends on how asset keys are structured)
                potential_key = f"{category}_{subcategory}_{asset}".replace(' ', '_')
                if potential_key.lower() == asset_key.lower():
                    return category
                # Also try without spaces or other variations
                if asset_key.lower().startswith(category.lower()):
                    return category
        
        # Fallback: try to extract from asset_key directly
        if '_' in asset_key:
            parts = asset_key.split('_')
            if parts[0].lower() in ['ground', 'space', 'link', 'user']:
                return parts[0]
        
        # Additional check for link assets (they might have different naming)
        asset_key_lower = asset_key.lower()
        if 'link' in asset_key_lower:
            return 'Link'
        
        return None
    
    def is_user_ground_segment_asset(self, asset_key):
        """Check if asset belongs to Ground -> User Ground Segment"""
        # Handle numerical asset keys from threat analysis (e.g., "1_probability", "2_probability")
        if '_probability' in asset_key:
            try:
                index = int(asset_key.split('_')[0]) - 1  # Convert to 0-based index
                if 0 <= index < len(self.ASSET_CATEGORIES):
                    category, subcategory, asset = self.ASSET_CATEGORIES[index]
                    return (category.lower() == 'ground' and 
                            subcategory.lower() == 'user ground segment')
            except (ValueError, IndexError):
                pass
        
        # Look for assets that are in Ground category and User Ground Segment subcategory
        for category_tuple in self.ASSET_CATEGORIES:
            if len(category_tuple) >= 3:
                category, subcategory, asset = category_tuple
                potential_key = f"{category}_{subcategory}_{asset}".replace(' ', '_')
                if (potential_key.lower() == asset_key.lower() and 
                    category.lower() == 'ground' and 
                    subcategory.lower() == 'user ground segment'):
                    return True
        
        # Also check if asset_key contains user ground segment indicators
        asset_key_lower = asset_key.lower()
        return ('user' in asset_key_lower and 'ground' in asset_key_lower)
    
    def normalize_threat_name(self, threat_name):
        """Normalize threat name for better matching"""
        # Remove common prefixes/suffixes and normalize
        normalized = threat_name.lower().strip()
        
        # Remove common threat prefixes
        prefixes_to_remove = ['malicious code / software / activity:', 'unauthorized modification:', 
                              'compromising confidential information (data breaches):']
        for prefix in prefixes_to_remove:
            if normalized.startswith(prefix):
                normalized = normalized[len(prefix):].strip()
        
        # Replace some common variations
        replacements = {
            'denial-of-service': 'denial of service',
            'dos': 'denial of service',
            'mitm': 'man-in-the-middle',
            'unauthorized access/hijacking': 'unauthorized access'
        }
        
        for old, new in replacements.items():
            if old in normalized:
                normalized = normalized.replace(old, new)
        
        return normalized
    
    def update_main_threat_table(self):
        """Update the main threat table with latest threat data"""
        # Use the existing method to update all threats in the main table
        if hasattr(self, 'update_all_threats_in_main_table'):
            self.update_all_threats_in_main_table()
        else:
            # Fallback method
            for threat in self.THREATS:
                if threat in self.threat_data and threat in self.threat_cells:
                    # Recalculate likelihood and impact for this threat
                    likelihood, impact, risk = self.calculate_threat_risk(threat)
                    
                    # Update table cells
                    cells = self.threat_cells[threat]
                    
                    # Update likelihood cell
                    if 'likelihood' in cells:
                        cells['likelihood'].config(text=likelihood, bg=self.get_risk_color(likelihood))
                    
                    # Update impact cell
                    if 'impact' in cells:
                        cells['impact'].config(text=impact, bg=self.get_risk_color(impact))
                    
                    # Update risk cell
                    if 'risk' in cells:
                        cells['risk'].config(text=risk, bg=self.get_risk_color(risk))
    
    def calculate_threat_risk(self, threat):
        """Calculate threat risk based on current data"""
        if threat not in self.threat_data:
            return "", "", ""
        
        asset_data = self.threat_data[threat]
        if not asset_data:
            return "", "", ""
        
        # Calculate average likelihood and impact across all assets
        total_likelihood = 0
        total_impact = 0
        asset_count = 0
        
        for asset_key, scores in asset_data.items():
            if isinstance(scores, list) and len(scores) >= 7:
                # Likelihood: average of first 5 criteria
                likelihood_scores = scores[:5]
                avg_likelihood = sum(likelihood_scores) / len(likelihood_scores)
                
                # Impact: average of last 2 criteria
                impact_scores = scores[5:7]
                avg_impact = sum(impact_scores) / len(impact_scores)
                
                total_likelihood += avg_likelihood
                total_impact += avg_impact
                asset_count += 1
        
        if asset_count == 0:
            return "", "", ""
        
        # Calculate final averages
        final_likelihood = total_likelihood / asset_count
        final_impact = total_impact / asset_count
        
        # Convert to risk levels
        likelihood_level = self.score_to_risk_level(final_likelihood)
        impact_level = self.score_to_risk_level(final_impact)
        
        # Calculate final risk using risk matrix
        risk_level = self.RISK_MATRIX.get((likelihood_level, impact_level), "Medium")
        
        return likelihood_level, impact_level, risk_level
    
    def score_to_risk_level(self, score):
        """Convert numeric score to risk level"""
        if score <= 1.5:
            return "Very Low"
        elif score <= 2.5:
            return "Low"
        elif score <= 3.5:
            return "Medium"
        elif score <= 4.5:
            return "High"
        else:
            return "Very High"
    
    def get_risk_color(self, risk_level):
        """Get color for risk level"""
        colors = {
            "Very Low": "#d4edda",
            "Low": "#d1ecf1", 
            "Medium": "#fff3cd",
            "High": "#f8d7da",
            "Very High": "#f5c6cb"
        }
        return colors.get(risk_level, "#ffffff")

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

        # Criteria table for THREATS
        self.create_criteria_table(content_frame, "threat")
        
        # Threat selection
        threat_frame = tk.Frame(content_frame, bg=self.COLORS['white'])
        threat_frame.pack(fill='x', pady=(20, 20))
        
        tk.Label(threat_frame, text="Select Threat:",
                font=('Segoe UI', 11, 'bold'),
                bg=self.COLORS['white'], fg=self.COLORS['dark']).pack(anchor='w')
        
        self.selected_threat_var = tk.StringVar()
        
        # Threat selection ComboBox is always enabled (even in read-only mode)
        threat_combo = ttk.Combobox(threat_frame,
                                   textvariable=self.selected_threat_var,
                                   values=self.THREATS,
                                   font=('Segoe UI', 10),
                                   state='readonly')
        threat_combo.pack(fill='x', pady=(5, 0))
        threat_combo.bind('<<ComboboxSelected>>', self.load_threat_data)
        
        # Disable mouse wheel on threat combobox
        self.disable_mousewheel_on_combobox(threat_combo)
        
        # Asset table for threat assessment
        self.create_assessment_table(content_frame, "threat")

        # Buttons frame
        buttons_frame = tk.Frame(content_frame, bg=self.COLORS['white'])
        buttons_frame.pack(pady=20)
        
        # Check if we're in read-only mode
        read_only_mode = getattr(self, 'threat_window_read_only', False)
        
        if read_only_mode:
            # Read-only mode: Show info message instead of save button
            info_label = tk.Label(buttons_frame, 
                                text="📖 READ-ONLY MODE - Assessment cannot be modified while controls are active",
                                font=('Segoe UI', 11, 'bold'),
                                bg='#f39c12', fg='white',
                                relief='flat', padx=25, pady=10)
            info_label.pack(side='left', padx=(0, 10))
        else:
            # Normal mode: Show save button
            save_btn = tk.Button(buttons_frame, text="SAVE ASSESSMENT",
                                font=('Segoe UI', 11, 'bold'),
                                bg=self.COLORS['success'], fg=self.COLORS['white'],
                                relief='flat', padx=25, pady=10,
                                command=lambda: self.save_threat_assessment(window))
            save_btn.pack(side='left', padx=(0, 10))
        
        # Help button (always available)
        help_btn = tk.Button(buttons_frame, text="❓ Help",
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['gray'], fg=self.COLORS['white'],
                            relief='flat', padx=20, pady=10,
                            command=self.show_help_threat)
        help_btn.pack(side='left')
              
        # Setup global mouse wheel scrolling for the entire content frame
        self.setup_global_mousewheel(content_frame, canvas)

    def show_help_threat(self):
        """Show help window with criteria descriptions"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Assessment Criteria - Help")
        help_window.geometry("1200x700")
        help_window.configure(bg=self.COLORS['white'])
        help_window.resizable(True, True)
        
        # Center the window
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Title
        title_label = tk.Label(help_window, text="Risk Assessment Criteria Descriptions - Threat Mode", 
                              font=('Segoe UI', 16, 'bold'),
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
            "Vulnerability effectiveness": "Assesses how effectively vulnerabilities can be exploited in the current system state.",
            "Mitigation Presence": "Evaluates the presence and effectiveness of security countermeasures.",
            "Detection Probability": "Measures the likelihood that malicious activities will be detected.",
            "Access Complexity": "Assesses how difficult it is for an attacker to gain access to the target.",
            "Privilege Requirement": "Evaluates the level of privileges needed to exploit the vulnerability.",
            "Response Delay": "Measures how quickly the organization can respond to security incidents.",
            "Resilience Impact": "Assesses the operational impact on system resilience and business continuity."
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
        
        # Add separator and tool explanation section
        separator_frame = tk.Frame(scrollable_frame, bg=self.COLORS['gray'], height=2)
        separator_frame.pack(fill='x', pady=(20, 15), padx=15)
        
        # Tool explanation title
        explanation_title = tk.Label(scrollable_frame, text="How the Risk Assessment Tool Works", 
                                    font=('Segoe UI', 14, 'bold'),
                                    bg=self.COLORS['white'], fg=self.COLORS['primary'])
        explanation_title.pack(pady=(10, 15), padx=15, anchor='w')
        
        # Tool explanation content
        explanation_frame = tk.Frame(scrollable_frame, bg=self.COLORS['light'], relief='ridge', bd=1)
        explanation_frame.pack(fill='x', padx=15, pady=(0, 20))
        
        explanation_text = """The Risk Assessment Tool for Phase B-C-D helps evaluate cybersecurity risks during detailed design and implementation phases of space missions. Here's how to use it effectively:

1. MISSION CONFIGURATION:
   • Configure mission parameters and security requirements for design/implementation phases
   • The tool adapts assessment criteria based on mission complexity and criticality
   • Load baseline data from previous Phase 0/A assessments for continuity

2. DETAILED THREAT ANALYSIS:
   • Click "THREAT ANALYSIS" to open the detailed assessment window
   • Select a specific threat from the dropdown menu to analyze
   • For each threat, evaluate all asset categories using 7 specific criteria on a scale of 1-5:
     - Vulnerability effectiveness, Mitigation Presence, Detection Probability, Access Complexity, Privilege Requirement, Response Delay, Resilience Impact
   • The tool automatically calculates Likelihood and Impact based on your assessments
   • Final Risk Level is determined using a standard risk matrix (Likelihood x Impact)

3. COMPREHENSIVE ASSET ANALYSIS:
   • Click "ASSET ANALYSIS" to open the asset-focused assessment window
   • Evaluate all assets using 9 detailed criteria covering both likelihood and impact factors
   • Asset criteria include: Dependency, Penetration, Cyber Maturity, Trust, Performance, Schedule, Costs, Reputation, Recovery
   • This provides a complementary view focusing on asset vulnerabilities and business impact

4. AUTOMATED RISK CALCULATIONS:
   • Advanced risk computation using multi-factor analysis: Risk = f(Threat, Vulnerability, Impact, Likelihood)
   • Dynamic risk scoring that adapts to mission phase and operational context
   • Risk aggregation across asset categories and threat domains
   • Confidence intervals and uncertainty analysis for risk estimates

5. DATA MANAGEMENT AND INTEGRATION (THREAT ANALYSIS MODE):
   • Use "Save Assessment" to temporarily store your current work in progress  
   • Use "Export Report" to generate final documentation and permanently save your analysis
   • IMPORTANT: "Save Assessment" stores data temporarily. For permanent documentation and final reports, always use "Export Report" to create Word/CSV documents
   • Import reports from Phase 0/A or external risk assessments using "IMPORT REPORT"
   • Import legacy data from previous 0-A reports using "IMPORT REPORT 0-A" (available in Output folder)
   • Maintain audit trails and version control for assessment iterations

6. HELP AND GUIDANCE:
   • Access context-sensitive help for both threat and asset analysis windows
   • Built-in guidance for industry-standard risk assessment methodologies
   • Reference materials and best practices for space mission security

7. RESULTS VISUALIZATION AND REPORTING:
   • Generate heat maps and risk matrices for stakeholder communication
   • Create executive dashboards with key risk indicators
   • Export findings in multiple formats for technical and management audiences
   • Produce compliance reports aligned with space industry standards

8. CONTINUOUS IMPROVEMENT:
   • Update assessments as design details and implementation plans evolve
   • Track risk mitigation effectiveness and residual risk levels
   • Integrate lessons learned and incident response feedback
   • Support iterative security design and validation processes"""
        
        explanation_label = tk.Label(explanation_frame, text=explanation_text,
                                   font=('Segoe UI', 10),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   anchor='nw', justify='left', wraplength=1150,
                                   padx=20, pady=15)
        explanation_label.pack(fill='both', expand=True)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        # Setup global mouse wheel scrolling for the help window
        self.setup_global_mousewheel(scrollable_frame, canvas)
        
        # Focus on help window
        help_window.focus_set()
        

    def create_asset_content(self, window):
        """Creates the asset content window (without threat selection)"""
        # Scrollable canvas with horizontal and vertical scrollbars
        outer_frame = tk.Frame(window, bg=self.COLORS['white'])
        outer_frame.pack(fill='both', expand=True, padx=5, pady=5)

        canvas = tk.Canvas(outer_frame, bg=self.COLORS['white'], highlightthickness=0)
        v_scrollbar = tk.Scrollbar(outer_frame, orient="vertical", command=canvas.yview)
        h_scrollbar = tk.Scrollbar(outer_frame, orient="horizontal", command=canvas.xview)
        canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        v_scrollbar.pack(side="right", fill="y")
        h_scrollbar.pack(side="bottom", fill="x")
        canvas.pack(side="left", fill="both", expand=True)

        content_frame = tk.Frame(canvas, bg=self.COLORS['white'])
        canvas.create_window((0, 0), window=content_frame, anchor="nw")
        content_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        # Criteria table for ASSETS
        self.create_asset_criteria_table(content_frame)
        # Asset table for asset assessment
        self.create_asset_assessment_table(content_frame)
        # Load latest asset data automatically
        self.load_latest_asset_data()

        # Buttons frame
        buttons_frame = tk.Frame(content_frame, bg=self.COLORS['white'])
        buttons_frame.pack(pady=10)
        save_btn = tk.Button(buttons_frame, text="SAVE ASSET ASSESSMENT",
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['success'], fg=self.COLORS['white'],
                            relief='flat', padx=15, pady=6,
                            command=lambda: self.save_asset_assessment(window))
        save_btn.pack(side='left', padx=(0, 8))
        help_btn = tk.Button(buttons_frame, text="❓ Help",
                            font=('Segoe UI', 11, 'bold'),
                            bg=self.COLORS['gray'], fg=self.COLORS['white'],
                            relief='flat', padx=12, pady=6,
                            command=self.show_help_asset)
        help_btn.pack(side='left')
        # Setup global mouse wheel scrolling for the entire content frame
        self.setup_global_mousewheel(content_frame, canvas)

    def show_help_asset(self):
        """Show help window with criteria descriptions"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Assessment Criteria - Help")
        help_window.geometry("1200x700")  # Asset analysis help window
        help_window.configure(bg=self.COLORS['white'])
        help_window.resizable(True, True)
        
        # Center the window
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Title
        title_label = tk.Label(help_window, text="Risk Assessment Criteria Descriptions - Asset Mode", 
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
            "Dependency": "Evaluates how critical the asset is to mission operations and business processes.",
            "Penetration": "Assesses the level of system access and control that can be gained through this asset.",
            "Cyber Maturity": "Evaluates the organization's cybersecurity governance and incident response capabilities.",
            "Trust": "Assesses the trustworthiness and security assurance of stakeholders involved with the asset.",
            "Performance": "Measures the impact on operational performance and service delivery capabilities.",
            "Schedule": "Evaluates the impact on project timelines and milestone achievement.",
            "Costs": "Assesses the financial impact and cost implications of security incidents.",
            "Reputation": "Evaluates the impact on organizational reputation and stakeholder confidence.",
            "Recovery": "Measures the time and effort required to restore normal operations after an incident."
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
        
        # Add separator and tool explanation section
        separator_frame = tk.Frame(scrollable_frame, bg=self.COLORS['gray'], height=2)
        separator_frame.pack(fill='x', pady=(20, 15), padx=15)
        
        # Tool explanation title
        explanation_title = tk.Label(scrollable_frame, text="How the Risk Assessment Tool Works", 
                                    font=('Segoe UI', 14, 'bold'),
                                    bg=self.COLORS['white'], fg=self.COLORS['primary'])
        explanation_title.pack(pady=(10, 15), padx=15, anchor='w')
        
        # Tool explanation content
        explanation_frame = tk.Frame(scrollable_frame, bg=self.COLORS['light'], relief='ridge', bd=1)
        explanation_frame.pack(fill='x', padx=15, pady=(0, 20))
        
        explanation_text = """The Risk Assessment Tool for Phase B-C-D helps evaluate cybersecurity risks during detailed design and implementation phases of space missions. Here's how to use it effectively:

1. MISSION CONFIGURATION:
   • Configure mission parameters and security requirements for design/implementation phases
   • The tool adapts assessment criteria based on mission complexity and criticality
   • Load baseline data from previous Phase 0/A assessments for continuity

2. DETAILED THREAT ANALYSIS:
   • Click "THREAT ANALYSIS" to open the detailed assessment window
   • Select a specific threat from the dropdown menu to analyze
   • For each threat, evaluate all asset categories using 7 specific criteria on a scale of 1-5:
     - Vulnerability effectiveness, Mitigation Presence, Detection Probability, Access Complexity, Privilege Requirement, Response Delay, Resilience Impact
   • The tool automatically calculates Likelihood and Impact based on your assessments
   • Final Risk Level is determined using a standard risk matrix (Likelihood x Impact)

3. COMPREHENSIVE ASSET ANALYSIS:
   • Click "ASSET ANALYSIS" to open the asset-focused assessment window
   • Evaluate all assets using 9 detailed criteria covering both likelihood and impact factors
   • Asset criteria include: Dependency, Penetration, Cyber Maturity, Trust, Performance, Schedule, Costs, Reputation, Recovery
   • This provides a complementary view focusing on asset vulnerabilities and business impact

4. AUTOMATED RISK CALCULATIONS:
   • Advanced risk computation using multi-factor analysis: Risk = f(Threat, Vulnerability, Impact, Likelihood)
   • Dynamic risk scoring that adapts to mission phase and operational context
   • Risk aggregation across asset categories and threat domains
   • Confidence intervals and uncertainty analysis for risk estimates

5. DATA MANAGEMENT AND INTEGRATION (ASSET ANALYSIS MODE):
   • Use "Save Assessment" to temporarily store your current work in progress  
   • Use "Export Report" to generate final documentation and permanently save your analysis
   • IMPORTANT: "Save Assessment" stores data temporarily. For permanent documentation and final reports, always use "Export Report" to create Word/CSV documents
   • Import reports from Phase 0/A or external risk assessments using "IMPORT REPORT"
   • Import legacy data from previous 0-A reports using "IMPORT REPORT 0-A" (available in Output folder)
   • Maintain audit trails and version control for assessment iterations

6. HELP AND GUIDANCE:
   • Access context-sensitive help for both threat and asset analysis windows
   • Built-in guidance for industry-standard risk assessment methodologies
   • Reference materials and best practices for space mission security

7. RESULTS VISUALIZATION AND REPORTING:
   • Generate heat maps and risk matrices for stakeholder communication
   • Create executive dashboards with key risk indicators
   • Export findings in multiple formats for technical and management audiences
   • Produce compliance reports aligned with space industry standards

8. CONTINUOUS IMPROVEMENT:
   • Update assessments as design details and implementation plans evolve
   • Track risk mitigation effectiveness and residual risk levels
   • Integrate lessons learned and incident response feedback
   • Support iterative security design and validation processes"""
        
        explanation_label = tk.Label(explanation_frame, text=explanation_text,
                                   font=('Segoe UI', 10),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   anchor='nw', justify='left', wraplength=1150,
                                   padx=20, pady=15)
        explanation_label.pack(fill='both', expand=True)
        
        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True, padx=(0, 5))
        scrollbar.pack(side="right", fill="y")
        
        # Setup global mouse wheel scrolling for the help window
        self.setup_global_mousewheel(scrollable_frame, canvas)
        
        # Focus on help window
        help_window.focus_set()
     

    def create_criteria_table(self, parent, assessment_type):
        """Creates the criteria table for threat or asset assessment"""
        title = f"{assessment_type.title()} Assessment Criteria"
        criteria_data = self.CRITERIA_DATA_THREAT if assessment_type == "threat" else self.CRITERIA_DATA_ASSET
        
        criteria_container = tk.LabelFrame(parent, 
                                         text=title,
                                         font=('Segoe UI', 12, 'bold'),
                                         bg=self.COLORS['white'], 
                                         fg=self.COLORS['primary'], 
                                         padx=20, 
                                         pady=15,
                                         relief='ridge', 
                                         bd=2)
        criteria_container.pack(fill='x', pady=(0, 20))

        # Celle della tabella criteri threat: prima colonna primary, header nero grassetto, colori trasposti
        for i, row in enumerate(criteria_data):
            for j, cell_text in enumerate(row):
                # Prima colonna: tutta primary, centrata, font size 10, grassetto
                if j == 0:
                    bg_color = self.COLORS['criteria_header']
                    fg_color = self.COLORS['white']
                    font_weight = 'bold'
                    font_size = 10
                    anchor = 'center'
                    justify = 'center'
                # Header: ogni cella colorata come la colonna, testo nero grassetto
                elif i == 0:
                    bg_color = self.CRITERIA_COLORS[(j-1) % len(self.CRITERIA_COLORS)]
                    fg_color = self.COLORS['dark']
                    font_weight = 'bold'
                    font_size = 10
                    anchor = 'center'
                    justify = 'center'
                # Celle dati: ogni colonna ha il suo colore
                else:
                    bg_color = self.CRITERIA_COLORS[(j-1) % len(self.CRITERIA_COLORS)]
                    fg_color = self.COLORS['dark']
                    font_weight = 'normal'
                    font_size = 8
                    anchor = 'nw'
                    justify = 'left'

                cell = tk.Label(
                    criteria_container, text=cell_text,
                    font=('Segoe UI', font_size, font_weight),
                    bg=bg_color,
                    fg=fg_color,
                    relief='ridge',
                    bd=1,
                    anchor=anchor,
                    justify=justify,
                    wraplength=180,
                    width=22,
                    height=3 if i == 0 else 4,
                    padx=3 if i == 0 else 6,
                    pady=2 if i == 0 else 3
                )
                cell.grid(row=i, column=j, padx=2, pady=2, sticky='ew', ipady=5)
        # Grid configuration with adjusted column sizes for transposed layout (Threat criteria)
        for j in range(8):  # Now we have 8 columns (Score + 7 criteria for threats)
            if j == 0:
                criteria_container.grid_columnconfigure(j, weight=1, minsize=40, uniform="criteria_cols")
            else:
                criteria_container.grid_columnconfigure(j, weight=1, minsize=80, uniform="criteria_cols")
        
        num_rows = len(criteria_data)
        for i in range(num_rows):
            criteria_container.grid_rowconfigure(i, minsize=70, uniform="criteria_rows")

    def create_asset_criteria_table(self, parent):
        """Creates the asset assessment criteria table"""
        criteria_container = tk.LabelFrame(parent, 
                                         text="Asset Assessment Criteria",
                                         font=('Segoe UI', 12, 'bold'),
                                         bg=self.COLORS['white'], 
                                         fg=self.COLORS['primary'], 
                                         padx=12, 
                                         pady=10,
                                         relief='ridge', 
                                         bd=2)
        criteria_container.pack(fill='x', pady=(0, 12))

        # Celle della tabella criteri asset: prima colonna primary, header nero grassetto, colori trasposti
        for i, row in enumerate(self.CRITERIA_DATA_ASSET):
            for j, cell_text in enumerate(row):
                # Prima colonna: tutta primary, centrata, font size 10, grassetto
                if j == 0:
                    bg_color = self.COLORS['criteria_header']
                    fg_color = self.COLORS['white']
                    font_weight = 'bold'
                    font_size = 10
                    anchor = 'center'
                    justify = 'center'
                # Header: ogni cella colorata come la colonna, testo nero grassetto
                elif i == 0:
                    bg_color = self.CRITERIA_COLORS[(j-1) % len(self.CRITERIA_COLORS)]
                    fg_color = self.COLORS['dark']
                    font_weight = 'bold'
                    font_size = 10
                    anchor = 'center'
                    justify = 'center'
                # Celle dati: ogni colonna ha il suo colore
                else:
                    bg_color = self.CRITERIA_COLORS[(j-1) % len(self.CRITERIA_COLORS)]
                    fg_color = self.COLORS['dark']
                    font_weight = 'normal'
                    font_size = 9
                    anchor = 'nw'
                    justify = 'left'

                cell = tk.Label(
                    criteria_container, text=cell_text,
                    font=('Segoe UI', font_size, font_weight),
                    bg=bg_color,
                    fg=fg_color,
                    relief='ridge',
                    bd=1,
                    anchor=anchor,
                    justify=justify,
                    wraplength=135,
                    width=16,
                    height=3 if i == 0 else 5,
                    padx=6,
                    pady=4
                )
                cell.grid(row=i, column=j, padx=2, pady=2, sticky='ew', ipady=6)
        for j in range(10):
            criteria_container.grid_columnconfigure(j, weight=1, minsize=60, uniform="criteria_cols")
        for i in range(6):
            criteria_container.grid_rowconfigure(i, minsize=48, uniform="criteria_rows")

    def create_assessment_table(self, parent, assessment_type):
        """Creates the assessment table for threat window only"""
        if assessment_type == "threat":
            self.create_threat_asset_table(parent)
        # Asset assessment is handled by separate function now

    def create_threat_asset_table(self, parent):
        """Creates the asset assessment table for threat window"""
        # Check if we're in read-only mode
        read_only_mode = getattr(self, 'threat_window_read_only', False)
        
        table_title = "Asset Assessment for Threat Analysis (Values 1-5)"
        if read_only_mode:
            table_title += " - READ ONLY MODE"
            
        table_frame = tk.LabelFrame(parent, text=table_title,
                                   font=('Segoe UI', 11, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=15, pady=15)
        table_frame.pack(fill='both', expand=True)
        
        # Headers for threat window - updated with all 7 threat criteria
        headers = ["Category", "Sub-category", "Component", "Vulnerability", "Mitigation", 
                  "Detection", "Access", "Privilege", "Response", 
                  "Resilience", "Likelihood", "Impact", "Risk"]
        
        for j, header in enumerate(headers):
            # Use different colors for criteria columns (3-9)
            if 3 <= j <= 9:  # Criteria columns
                bg_color = self.CRITERIA_COLORS[(j-3) % len(self.CRITERIA_COLORS)]
                fg_color = self.COLORS['dark']
            else:  # Standard columns (Category, Sub-category, Component, Likelihood, Impact, Risk)
                bg_color = self.COLORS['primary']
                fg_color = self.COLORS['white']
                
            cell = tk.Label(table_frame, text=header,
                           font=('Segoe UI', 9, 'bold'),
                           bg=bg_color, fg=fg_color,
                           relief='ridge', bd=1, width=8,
                           wraplength=80)
            cell.grid(row=0, column=j, padx=1, pady=1, sticky='ew')

        # Reset threat-specific variables
        self.threat_combo_vars = {}
        self.threat_impact_entries = {}
        self.combo_vars = self.threat_combo_vars
        self.impact_entries = self.threat_impact_entries
        
        # Asset rows
        for i in range(len(self.ASSET_CATEGORIES)):
            category, sub_category, component = self.ASSET_CATEGORIES[i]
            asset_key = f"{i+1}_probability"

            # Category (read-only)
            cat_cell = tk.Label(table_frame, text=category,
                               font=('Segoe UI', 9, 'bold'),
                               bg=self.COLORS['light'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, width=15,
                               wraplength=70)
            cat_cell.grid(row=i+1, column=0, padx=1, pady=1, sticky='ew')
            
            # Sub-category (read-only)
            sub_cat_cell = tk.Label(table_frame, text=sub_category,
                                   font=('Segoe UI', 9),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   relief='ridge', bd=1, width=18,
                                   wraplength=110)
            sub_cat_cell.grid(row=i+1, column=1, padx=1, pady=1, sticky='ew')
            
            # Component (read-only)
            comp_cell = tk.Label(table_frame, text=component,
                                font=('Segoe UI', 9),
                                bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                relief='ridge', bd=1, width=25,
                                wraplength=150)
            comp_cell.grid(row=i+1, column=2, padx=1, pady=1, sticky='ew')

            # Storage for this row
            row_entries = {}
            self.combo_vars[asset_key] = {}

            # Editable columns (3-9: All 7 threat criteria)
            for j in range(3, 10):
                combo_var = tk.StringVar(value="")
                # Use custom style for each criterion
                style_name = f"Criteria{j-3}.TCombobox"
                
                # Check if we're in read-only mode for threat window
                read_only_mode = getattr(self, 'threat_window_read_only', False)
                combo_state = 'disabled' if read_only_mode else 'readonly'
                
                combo = ttk.Combobox(table_frame,
                                    textvariable=combo_var,
                                    values=["", "1", "2", "3", "4", "5"],
                                    font=('Segoe UI', 8),
                                    width=5, state=combo_state,
                                    style=style_name)
                combo.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                
                # Disable mouse wheel on combobox
                self.disable_mousewheel_on_combobox(combo)
                
                row_entries[j-3] = combo
                self.combo_vars[asset_key][j-3] = combo_var

                # Bind calculations for threat context
                if j <= 7:  # First 5 criteria (Vulnerability, Mitigation, Detection, Access, Privilege) -> Likelihood
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.safe_calculate_likelihood(key))
                elif j <= 9:  # Last 2 criteria (Response Delay, Resilience Impact) -> Impact
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.safe_calculate_impact(key))
            
            # Calculated columns (10-12: Likelihood, Impact, Risk) - read-only
            for j in range(10, 13):
                calc_cell = tk.Label(table_frame, text="",
                                   font=('Segoe UI', 8),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   relief='ridge', bd=1, width=8)
                calc_cell.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                row_entries[j-3] = calc_cell
            
            self.impact_entries[asset_key] = row_entries
        
        # Grid configuration
        table_frame.grid_columnconfigure(0, weight=1, minsize=80, uniform="category_cols")
        table_frame.grid_columnconfigure(1, weight=1, minsize=120, uniform="sub_category_col")
        table_frame.grid_columnconfigure(2, weight=1, minsize=140, uniform="component_col")
        for j in range(3, 10):  # 7 threat criteria columns
            table_frame.grid_columnconfigure(j, weight=1, minsize=90, uniform="criteria_cols")
        for j in range(10, 13):  # 3 calculated columns (Likelihood, Impact, Risk)
            table_frame.grid_columnconfigure(j, weight=1, minsize=80, uniform="calc_cols")
        
        for i in range(len(self.ASSET_CATEGORIES) + 1):
            table_frame.grid_rowconfigure(i, minsize=40, uniform="rows")

        # Add color legend for threat criteria
        self.create_threat_color_legend(parent)

    def create_threat_color_legend(self, parent):
        """Creates a color legend for threat criteria"""
        threat_criteria = ["Vulnerability", "Mitigation", "Detection", "Access", 
                          "Privilege", "Response", "Resilience"]
        self.create_color_legend(parent, threat_criteria)

    def create_asset_assessment_table(self, parent):
        """Creates the asset assessment table for asset window"""
        table_frame = tk.LabelFrame(parent, text="Asset Assessment (Values 1-5)",
                                   font=('Segoe UI', 11, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=15, pady=15)
        table_frame.pack(fill='both', expand=True)

        # Headers for asset window - includes all 9 criteria columns (no Risk column)
        headers = ["Category", "Sub-category", "Component", "Dependency", "Penetration", "Maturity", "Trust",
                  "Performance", "Schedule", "Costs", "Reputation", "Recovery", "Likelihood", "Impact"]
        
        for j, header in enumerate(headers):
            # Use different colors for criteria columns (3-11)
            if 3 <= j <= 11:  # Criteria columns
                bg_color = self.CRITERIA_COLORS[(j-3) % len(self.CRITERIA_COLORS)]
                fg_color = self.COLORS['dark']
            else:  # Standard columns (Category, Sub-category, Component, Likelihood, Impact)
                bg_color = self.COLORS['primary']
                fg_color = self.COLORS['white']
                
            cell = tk.Label(table_frame, text=header,
                           font=('Segoe UI', 9, 'bold'),
                           bg=bg_color, fg=fg_color,
                           relief='ridge', bd=1, width=10,
                           wraplength=100)
            cell.grid(row=0, column=j, padx=1, pady=1, sticky='ew')

        # Reset asset-specific variables
        self.asset_combo_vars = {}
        self.asset_impact_entries = {}
        self.combo_vars = self.asset_combo_vars
        self.impact_entries = self.asset_impact_entries
        
        # Asset rows
        for i in range(len(self.ASSET_CATEGORIES)):
            category, sub_category, component = self.ASSET_CATEGORIES[i]
            asset_key = f"{i+1}_probability"

            # Category (read-only)
            cat_cell = tk.Label(table_frame, text=category,
                               font=('Segoe UI', 8, 'bold'),
                               bg=self.COLORS['light'], fg=self.COLORS['dark'],
                               relief='ridge', bd=1, width=12,
                               wraplength=70)
            cat_cell.grid(row=i+1, column=0, padx=1, pady=1, sticky='ew')
            
            # Sub-category (read-only)
            sub_cat_cell = tk.Label(table_frame, text=sub_category,
                                   font=('Segoe UI', 8),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   relief='ridge', bd=1, width=15,
                                   wraplength=110)
            sub_cat_cell.grid(row=i+1, column=1, padx=1, pady=1, sticky='ew')
            
            # Component (read-only)
            comp_cell = tk.Label(table_frame, text=component,
                                font=('Segoe UI', 8),
                                bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                relief='ridge', bd=1, width=18,
                                wraplength=80)
            comp_cell.grid(row=i+1, column=2, padx=1, pady=1, sticky='ew')

            # Storage for this row
            row_entries = {}
            self.combo_vars[asset_key] = {}

            # Editable columns (3-11: All 9 criteria)
            for j in range(3, 12):
                combo_var = tk.StringVar(value="")
                # Use custom style for each criterion
                style_name = f"Criteria{j-3}.TCombobox"
                combo = ttk.Combobox(table_frame,
                                    textvariable=combo_var,
                                    values=["", "1", "2", "3", "4", "5"],
                                    font=('Segoe UI', 7),
                                    width=4, state='readonly',
                                    style=style_name)
                combo.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                
                # Disable mouse wheel on combobox
                self.disable_mousewheel_on_combobox(combo)
                
                row_entries[j-3] = combo
                self.combo_vars[asset_key][j-3] = combo_var

                # Bind calculations for asset context
                if j <= 6:  # First 4 criteria (Dependency, Penetration, Cyber Maturity, Trust) -> Likelihood
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.safe_calculate_likelihood(key))
                elif j <= 11:  # Other 5 criteria (Performance, Schedule, Costs, Reputation, Recovery) -> Impact
                    combo_var.trace_add('write', lambda *args, key=asset_key: self.safe_calculate_impact(key))

            # Calculated columns (12-13: Likelihood, Impact) - read-only, no Risk column
            for j in range(12, 14):
                calc_cell = tk.Label(table_frame, text="",
                                   font=('Segoe UI', 8),
                                   bg=self.COLORS['light'], fg=self.COLORS['dark'],
                                   relief='ridge', bd=1, width=8)
                calc_cell.grid(row=i+1, column=j, padx=1, pady=1, sticky='ew')
                row_entries[j-3] = calc_cell
            
            self.impact_entries[asset_key] = row_entries
        
        # Grid configuration
        table_frame.grid_columnconfigure(0, weight=1, minsize=80, uniform="category_cols")
        table_frame.grid_columnconfigure(1, weight=1, minsize=120, uniform="sub_category_col")
        table_frame.grid_columnconfigure(2, weight=1, minsize=140, uniform="component_col")
        for j in range(3, 12):  # 9 asset criteria columns
            table_frame.grid_columnconfigure(j, weight=1, minsize=70, uniform="criteria_cols")
        for j in range(12, 14):  # 2 calculated columns (Likelihood, Impact) - no Risk
            table_frame.grid_columnconfigure(j, weight=1, minsize=80, uniform="calc_cols")
        
        for i in range(len(self.ASSET_CATEGORIES) + 1):
            table_frame.grid_rowconfigure(i, minsize=40, uniform="rows")

        # Add color legend below the table
        self.create_asset_color_legend(parent)

    def create_asset_color_legend(self, parent):
        """Creates a color legend for asset criteria"""
        asset_criteria = ["Dependency", "Penetration", "Maturity", "Trust", 
                         "Performance", "Schedule", "Costs", "Reputation", "Recovery"]
        self.create_color_legend(parent, asset_criteria)

    def create_color_legend(self, parent, criteria_names):
        """Creates a color legend showing which color corresponds to which criterion - single row layout"""
        legend_frame = tk.LabelFrame(parent, text="Criteria Color Legend",
                                   font=('Segoe UI', 10, 'bold'),
                                   bg=self.COLORS['white'], fg=self.COLORS['primary'],
                                   padx=5, pady=5)
        legend_frame.pack(fill='x', pady=(10, 0))
        
        # Single row container for all legend items
        legend_container = tk.Frame(legend_frame, bg=self.COLORS['white'])
        legend_container.pack(expand=True)
        
        # Create legend entries in a single row layout
        for i, criterion in enumerate(criteria_names):
            # Container for each criterion (color + text)
            criterion_frame = tk.Frame(legend_container, bg=self.COLORS['white'])
            criterion_frame.pack(side='left', padx=3)
            
            # Color square
            color_square = tk.Label(criterion_frame, text="  ", 
                                  bg=self.CRITERIA_COLORS[i],
                                  relief='ridge', bd=1, width=2, height=1)
            color_square.pack(side='left', padx=(0, 2))
            
            # Criterion name
            name_label = tk.Label(criterion_frame, text=criterion,
                                font=('Segoe UI', 7),
                                bg=self.COLORS['white'], fg=self.COLORS['dark'])
            name_label.pack(side='left')    # ===== CALCULATION METHODS =====
    
    def safe_calculate_likelihood(self, key):
        """Safely calculates likelihood with error handling - context aware"""
        try:
            if not self.validate_combo_values(key):
                # Use correct column index based on context: threat window=7, asset window=9
                col_idx = 7 if hasattr(self, 'selected_threat_var') and self.combo_vars is self.threat_combo_vars else 9
                self.update_display(key, col_idx, "")
                return
            
            # Use appropriate calculation method based on context
            if self.combo_vars is self.threat_combo_vars:
                # We're in threat window
                self.calculate_threat_likelihood(key)
            else:
                # We're in asset window
                self.calculate_asset_likelihood(key)
                
        except Exception as e:
            logging.error(f"Error calculating likelihood for {key}: {e}")
            col_idx = 7 if self.combo_vars is self.threat_combo_vars else 9
            self.update_display(key, col_idx, "")
            messagebox.showerror("Calculation Error", f"Error calculating likelihood: {str(e)}")

    def safe_calculate_impact(self, key):
        """Safely calculates impact with error handling - context aware"""
        try:
            if not self.validate_combo_values(key):
                # Use correct column index based on context: threat window=8, asset window=10
                col_idx = 8 if self.combo_vars is self.threat_combo_vars else 10
                self.update_display(key, col_idx, "")
                return
            
            # Use appropriate calculation method based on context
            if self.combo_vars is self.threat_combo_vars:
                # We're in threat window
                self.calculate_threat_impact(key)
            else:
                # We're in asset window
                self.calculate_asset_impact(key)
                
        except Exception as e:
            logging.error(f"Error calculating impact for {key}: {e}")
            col_idx = 8 if self.combo_vars is self.threat_combo_vars else 10
            self.update_display(key, col_idx, "")
            messagebox.showerror("Calculation Error", f"Error calculating impact: {str(e)}")

    def calculate_threat_likelihood(self, key):
        """Calculates Threat Likelihood combining threat-specific likelihood with asset likelihood"""
        if key not in self.combo_vars:
            return

        # Calculate threat-specific likelihood from first 5 criteria (columns 0-4)
        threat_values = []
        for col_idx in [0, 1, 2, 3, 4]:
            if col_idx not in self.combo_vars[key]:
                continue
                self.update_display(key, 7, "")
                return
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                continue
                self.update_display(key, 7, "")
                return
            try:
                threat_values.append(float(value_str))
            except ValueError:
                self.update_display(key, 7, "")
                return
        if len(threat_values) == 0:
            self.update_display(key, 7, "")
            self.update_display(key, 9, "")
            return
        
        if len(threat_values) > 0:
            # Calculate threat-specific likelihood using quadratic mean
            threat_quadratic_mean = math.sqrt(sum(x**2 for x in threat_values) / len(threat_values))
            threat_likelihood = (threat_quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            threat_likelihood = max(0.0, min(1.0, threat_likelihood))
            
            # Get asset likelihood from latest asset assessment
            asset_likelihood = self.get_asset_likelihood_for_key(key)
            
            if asset_likelihood >= 0:
                # Convert both likelihoods to categories
                threat_likelihood_cat = self.value_to_category(threat_likelihood)
                asset_likelihood_cat = self.value_to_category(asset_likelihood)
                #print (key)
                # Combine using ISO 27005 risk matrix (treat as likelihood x likelihood)
                combined_likelihood_cat = self.RISK_MATRIX.get((threat_likelihood_cat, asset_likelihood_cat), threat_likelihood_cat)
                
                self.update_display(key, 7, combined_likelihood_cat)
                
                # Recalculate risk if Impact is available
                self.calculate_risk(key)
            else:
                # If no asset data available, result must be empty
                self.update_display(key, 7, "")
                self.calculate_risk(key)

    def calculate_threat_impact(self, key):
        """Calculates Threat Impact using last 2 criteria (quadratic mean)"""
        if key not in self.combo_vars:
            return

        # Get values for last 2 threat criteria (columns 5,6)
        values = []
        for col_idx in [5, 6]:
            if col_idx not in self.combo_vars[key]:
                continue
                self.update_display(key, 8, "")
                return
            
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                continue
                self.update_display(key, 8, "")
                return
            
            try:
                values.append(float(value_str))
            except ValueError:
                self.update_display(key, 8, "")
                return
        
        if len(values) == 0:
            self.update_display(key, 8, "")
            self.update_display(key, 9, "")
            return

        if len(values) >0 :
            # Calculate threat-specific impact using quadratic mean
            threat_quadratic_mean = math.sqrt(sum(x**2 for x in values) / len(values))
            threat_impact = (threat_quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            threat_impact = max(0.0, min(1.0, threat_impact))
            
            # Get asset impact from latest asset assessment
            asset_impact = self.get_asset_impact_for_key(key)
            
            if asset_impact >= 0:
                # Convert both impacts to categories
                threat_impact_cat = self.value_to_category(threat_impact)
                asset_impact_cat = self.value_to_category(asset_impact)
                
                # Combine using ISO 27005 risk matrix (treat as impact x impact)
                combined_impact_cat = self.RISK_MATRIX.get((threat_impact_cat, asset_impact_cat), threat_impact_cat)
                
                self.update_display(key, 8, combined_impact_cat)
                
                # Recalculate risk
                self.calculate_risk(key)
            else:
                # If no asset data available, result must be empty
                self.update_display(key, 8, "")
                self.calculate_risk(key)

    def calculate_asset_likelihood(self, key):
        """Calculates Asset Likelihood using quadratic mean (First 4 criteria: Dependency, Penetration, Cyber Maturity, Trust)"""
        if key not in self.combo_vars:
            return

        # Get values for first 4 criteria (columns 0,1,2,3)
        values = []
        for col_idx in [0, 1, 2, 3]:
            if col_idx not in self.combo_vars[key]:
                continue
                self.update_display(key, 9, "")  # Likelihood is at column 9
                return
            
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                continue
                self.update_display(key, 9, "")
                return
            
            try:
                values.append(float(value_str))
            except ValueError:
                self.update_display(key, 9, "")
                return
        if len(values) == 0:
            self.update_display(key, 9, "")
            return
        
        if len(values) >0:
            # Use quadratic mean for likelihood calculation
            quadratic_mean = math.sqrt(sum(x**2 for x in values) / len(values))
            likelihood = (quadratic_mean - 1) / 4  # [1,5] -> [0,1]
            likelihood = max(0.0, min(1.0, likelihood))            # Convert to category
            likelihood_category = self.value_to_category(likelihood)
            self.update_display(key, 9, likelihood_category)

            # Recalculate risk
            self.calculate_risk(key)

    def calculate_asset_impact(self, key):
        """Calculates Asset Impact using quadratic mean (Last 5 criteria: Performance, Schedule, Costs, Reputation, Recovery)"""
        if key not in self.combo_vars:
            return        # Get values for last 5 criteria (columns 4,5,6,7,8)
        values = []
        for col_idx in [4, 5, 6, 7, 8]:
            if col_idx not in self.combo_vars[key]:
                continue
                self.update_display(key, 10, "")  # Impact is at column 13
                return
            
            value_str = self.combo_vars[key][col_idx].get().strip()
            if not value_str or value_str == "0":
                continue
                self.update_display(key, 10, "")
                return
            try:
                values.append(float(value_str))
            except ValueError:
                self.update_display(key, 10, "")
                return
        if len(values) == 0:
            self.update_display(key, 10, "")
            return
        
        if len(values) > 0:
            # For assets, use quadratic mean for more conservative approach
            quadratic_mean = math.sqrt(sum(x**2 for x in values) / len(values))
            impact = (quadratic_mean - 1) / 4  # [1,5] -> [0,1]
            impact = max(0.0, min(1.0, impact))            # Convert to category
            impact_category = self.value_to_category(impact)
            self.update_display(key, 10, impact_category)            # Recalculate risk
            self.calculate_risk(key)

    def calculate_risk(self, key):
        """Calculates Risk using the Likelihood x Impact matrix"""
        if key not in self.impact_entries:
            return
        
        # Only calculate risk for threat window (asset window has no Risk column)
        if self.combo_vars is not self.threat_combo_vars:
            return
        
        # Threat window column indices
        likelihood_idx, impact_idx, risk_idx = 7, 8, 9
        
        # Get Likelihood and Impact (categories)
        likelihood_widget = self.impact_entries[key][likelihood_idx]
        impact_widget = self.impact_entries[key][impact_idx]
        
        likelihood_cat = likelihood_widget.cget('text') if hasattr(likelihood_widget, 'cget') else ""
        impact_cat = impact_widget.cget('text') if hasattr(impact_widget, 'cget') else ""
        
        if not likelihood_cat or not impact_cat:
            self.update_display(key, risk_idx, "")
            return

        # Check if they are valid categories
        valid_categories = ["Very Low", "Low", "Medium", "High", "Very High"]
        if likelihood_cat in valid_categories and impact_cat in valid_categories:
            # Get risk from matrix
            risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
            self.update_display(key, risk_idx, risk_level)
        else:
            self.update_display(key, risk_idx, "")

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

    def validate_combo_values(self, key):
        """Validates that combo box values are in correct range"""
        if key not in self.combo_vars:
            return False
        
        for col_idx, combo_var in self.combo_vars[key].items():
            value = combo_var.get().strip()
            if value and value not in ["1", "2", "3", "4", "5"]:
                logging.warning(f"Invalid value '{value}' found in asset {key}, column {col_idx}")
                combo_var.set("")
                return False
        return True

    # ===== DATA MANAGEMENT =====

    def load_threat_data(self, event=None):
        """Load data for selected threat and update GUI"""
        selected_threat = self.selected_threat_var.get()
        
        # Clear all data first
        for key in self.impact_entries:
            self.clear_data(key)
        
        # Load saved data if exists
        if selected_threat and selected_threat in self.threat_data:
            threat_data = self.threat_data[selected_threat]
            for key, row_data in threat_data.items():
                if key in self.combo_vars:
                    for col_idx, value in row_data.items():
                        if int(col_idx) in self.combo_vars[key]:
                            self.combo_vars[key][int(col_idx)].set(value)
        
        # Recalculate all values
        for key in self.impact_entries:
            self.safe_calculate_likelihood(key)
            self.safe_calculate_impact(key)

    def load_latest_asset_data(self):
        """Automatically load the latest saved asset data"""
        if not self.asset_data:
            return
        
        # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
        assessment_keys = [key for key in self.asset_data.keys() if key.startswith('assessment_')]
        imported_keys = [key for key in self.asset_data.keys() if key.startswith('imported_')]
        
        # Use the latest assessment key if available, otherwise use latest imported key
        if assessment_keys:
            latest_key = max(assessment_keys)
        elif imported_keys:
            latest_key = max(imported_keys)
        else:
            latest_key = max(self.asset_data.keys()) if self.asset_data else None
        
        if latest_key and latest_key in self.asset_data:
            asset_data = self.asset_data[latest_key]
            
            # Load data into comboboxes
            for key, row_data in asset_data.items():
                if key in self.combo_vars:
                    for col_idx, value in row_data.items():
                        col_index = int(col_idx)
                        if col_index in self.combo_vars[key]:
                            self.combo_vars[key][col_index].set(value)
            
            # Recalculate everything after loading data
            for key in self.impact_entries:
                self.safe_calculate_likelihood(key)
                self.safe_calculate_impact(key)

    def clear_data(self, key):
        """Clear data for a row"""
        if key in self.combo_vars:
            for combo_var in self.combo_vars[key].values():
                combo_var.set("")
        
        if key in self.impact_entries:
            if self.selected_threat_var:  # Threat window
                indices = [7, 8, 9]
            else:  # Asset window
                indices = [9, 10]
            
            for col_idx in indices:
                if col_idx in self.impact_entries[key]:
                    self.impact_entries[key][col_idx].config(text="")

    def save_threat_assessment(self, window):
        """Save threat assessment data"""
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
        
        # Update main table
        self.update_all_threats_in_main_table()
        
        messagebox.showinfo("Success", f"Assessment for '{selected_threat}' saved!")
        window.destroy()

    def save_asset_assessment(self, window):
        """Save asset assessment data with timestamp"""
        asset_data = {}
        for key in self.combo_vars:
            row_data = {}
            for col_idx, combo_var in self.combo_vars[key].items():
                value = combo_var.get().strip()
                if value:
                    row_data[str(col_idx)] = value
            if row_data:
                asset_data[key] = row_data
        
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.asset_data[f"assessment_{timestamp}"] = asset_data
        
        # Update main table since asset values affect threat calculations
        self.update_all_threats_in_main_table()
        
        messagebox.showinfo("Success", f"Asset assessment saved successfully!\n{len(asset_data)} assets evaluated.")
        window.destroy()

    def update_all_threats_in_main_table(self):
        """Update main table with the likelihood, impact and risk that produce the maximum risk for each threat"""
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}

        for threat_name in self.threat_data.keys():
            if threat_name not in self.threat_cells:
                continue

            # Updata main table
            likelihood, impact, risk = self.get_max_risk_combination(self.threat_data[threat_name])
            self.threat_cells[threat_name]['likelihood'].config(text=likelihood)
            self.threat_cells[threat_name]['impact'].config(text=impact)
            self.threat_cells[threat_name]['risk'].config(text=risk)

    def calculate_likelihood_from_saved_data(self, asset_data):
        """Calculate likelihood from saved data for threats combining threat and asset likelihood"""
        try:
            # Calculate threat-specific likelihood from first 5 criteria
            threat_values = []
            for i in [0, 1, 2, 3, 4]:
                # Use get() with default to safely access the value
                val = asset_data.get(str(i), "")
                # Skip empty or zero values but continue with remaining criteria
                if val and str(val).strip() and str(val) != "0":
                    try:
                        threat_values.append(float(val))
                    except (ValueError, TypeError):
                        continue  # Skip invalid values but continue processing
            
            # Require at least one valid value to calculate threat likelihood
            if not threat_values:
                return -1.0
            
            # Calculate threat-specific likelihood using quadratic mean
            threat_quadratic_mean = math.sqrt(sum(x**2 for x in threat_values) / len(threat_values))
            threat_likelihood = (threat_quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            threat_likelihood = max(0.0, min(1.0, threat_likelihood))
            
            # Get asset likelihood from asset assessment
            asset_likelihood = -1.0
            if hasattr(self, 'asset_data') and self.asset_data:
                # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
                assessment_keys = [key for key in self.asset_data.keys() if key.startswith('assessment_')]
                imported_keys = [key for key in self.asset_data.keys() if key.startswith('imported_')]
                
                # Use the latest assessment key if available, otherwise use latest imported key
                if assessment_keys:
                    latest_key = max(assessment_keys)
                elif imported_keys:
                    latest_key = max(imported_keys)
                else:
                    latest_key = max(self.asset_data.keys()) if self.asset_data else None
                
                if latest_key and latest_key in self.asset_data:
                    # Try to find matching asset data
                    for asset_key, asset_assessment_data in self.asset_data[latest_key].items():
                        # Calculate asset likelihood for comparison
                        asset_values = []
                        for i in [0, 1, 2, 3]:
                            val = asset_assessment_data.get(str(i), "")
                            if val and str(val).strip() and str(val) != "0":
                                try:
                                    asset_values.append(float(val))
                                except (ValueError, TypeError):
                                    continue
                        
                        if asset_values:  # If we have at least one valid value
                            asset_quadratic_mean = math.sqrt(sum(x**2 for x in asset_values) / len(asset_values))
                            asset_likelihood = (asset_quadratic_mean - 1) / 4
                            asset_likelihood = max(0.0, min(1.0, asset_likelihood))
                            break  # Use first valid asset likelihood found
            
            # Combine threat and asset likelihood if asset data is available
            if asset_likelihood >= 0:
                # Convert both likelihoods to categories
                threat_likelihood_cat = self.value_to_category(threat_likelihood)
                asset_likelihood_cat = self.value_to_category(asset_likelihood)
                
                # Combine using ISO 27005 risk matrix (treat as likelihood x likelihood)
                combined_likelihood_cat = self.RISK_MATRIX.get((threat_likelihood_cat, asset_likelihood_cat), threat_likelihood_cat)
                
                # Convert back to numeric value for consistency with return type
                category_to_value = {
                    "Very Low": 0.05,
                    "Low": 0.25, 
                    "Medium": 0.55,
                    "High": 0.8,
                    "Very High": 0.95
                }
                return category_to_value.get(combined_likelihood_cat, threat_likelihood)
            else:
                # If no asset data, return threat likelihood alone
                return threat_likelihood
                
        except Exception as e:
            return -1.0

    def calculate_impact_from_saved_data(self, asset_data):
        """Calculate impact from saved data for threats combining threat and asset impact"""
        try:
            # Calculate threat-specific impact from last 2 criteria
            threat_values = []
            for i in [5, 6]:
                # Use get() with default to safely access the value
                val = asset_data.get(str(i), "")
                # Skip empty or zero values but continue with remaining criteria
                if val and str(val).strip() and str(val) != "0":
                    try:
                        threat_values.append(float(val))
                    except (ValueError, TypeError):
                        continue  # Skip invalid values but continue processing
            
            # Require at least one valid value to calculate threat impact
            if not threat_values:
                return -1.0
            
            # Calculate threat-specific impact using quadratic mean
            threat_quadratic_mean = math.sqrt(sum(x**2 for x in threat_values) / len(threat_values))
            threat_impact = (threat_quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            threat_impact = max(0.0, min(1.0, threat_impact))
            
            # Get asset impact from asset assessment
            asset_impact = -1.0
            if hasattr(self, 'asset_data') and self.asset_data:
                # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
                assessment_keys = [key for key in self.asset_data.keys() if key.startswith('assessment_')]
                imported_keys = [key for key in self.asset_data.keys() if key.startswith('imported_')]
                
                # Use the latest assessment key if available, otherwise use latest imported key
                if assessment_keys:
                    latest_key = max(assessment_keys)
                elif imported_keys:
                    latest_key = max(imported_keys)
                else:
                    latest_key = max(self.asset_data.keys()) if self.asset_data else None
                
                if latest_key and latest_key in self.asset_data:
                    # Try to find matching asset data
                    for asset_key, asset_assessment_data in self.asset_data[latest_key].items():
                        # Calculate asset impact for comparison
                        asset_values = []
                        for i in [4, 5, 6, 7, 8]:
                            val = asset_assessment_data.get(str(i), "")
                            if val and str(val).strip() and str(val) != "0":
                                try:
                                    asset_values.append(float(val))
                                except (ValueError, TypeError):
                                    continue
                        
                        if asset_values:  # If we have at least one valid value
                            asset_quadratic_mean = math.sqrt(sum(x**2 for x in asset_values) / len(asset_values))
                            asset_impact = (asset_quadratic_mean - 1) / 4
                            asset_impact = max(0.0, min(1.0, asset_impact))
                            break  # Use first valid asset impact found
            
            # Combine threat and asset impact if asset data is available
            if asset_impact >= 0:
                # Convert both impacts to categories
                threat_impact_cat = self.value_to_category(threat_impact)
                asset_impact_cat = self.value_to_category(asset_impact)
                
                # Combine using ISO 27005 risk matrix (treat as impact x impact)
                combined_impact_cat = self.RISK_MATRIX.get((threat_impact_cat, asset_impact_cat), threat_impact_cat)
                
                # Convert back to numeric value for consistency with return type
                category_to_value = {
                    "Very Low": 0.05,
                    "Low": 0.25, 
                    "Medium": 0.55,
                    "High": 0.8,
                    "Very High": 0.95
                }
                return category_to_value.get(combined_impact_cat, threat_impact)
            else:
                # If no asset data, return threat impact alone
                return threat_impact
                
        except Exception as e:
            return -1.0

    def get_max_risk_combination(self, threat_data):
        """
        Restituisce (likelihood_cat, impact_cat, risk_cat) dell'asset che ha il rischio massimo per un threat.
        threat_data: dict delle righe asset per uno specifico threat (es: self.threat_data[threat_name])
        """
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}
        best_likelihood = ""
        best_impact = ""
        best_risk = ""
        max_priority = 0

        # Trova la chiave dell'ultimo asset assessment
        latest_key = None
        if hasattr(self, 'asset_data') and self.asset_data:
            assessment_keys = [k for k in self.asset_data.keys() if k.startswith('assessment_')]
            imported_keys = [k for k in self.asset_data.keys() if k.startswith('imported_')]
            if assessment_keys:
                latest_key = max(assessment_keys)
            elif imported_keys:
                latest_key = max(imported_keys)
            elif self.asset_data:
                latest_key = max(self.asset_data.keys())

        asset_assessment = self.asset_data[latest_key] if latest_key and latest_key in self.asset_data else {}

        for asset_key, asset_data in threat_data.items():
            # --- Likelihood ---
            # Threat-specific likelihood (primi 5 criteri)
            threat_values = []
            for i in [0, 1, 2, 3, 4]:
                val = asset_data.get(str(i), "")
                if val and str(val) != "0":
                    try:
                        threat_values.append(float(val))
                    except (ValueError, TypeError):
                        pass
            
            if not threat_values:  # Se non ci sono valori validi, skip
                continue
                
            threat_quadratic_mean = math.sqrt(sum(x**2 for x in threat_values) / len(threat_values))
            threat_likelihood = (threat_quadratic_mean - 1) / 4
            threat_likelihood = max(0.0, min(1.0, threat_likelihood))

            # Asset likelihood (primi 4 criteri)
            asset_likelihood = -1.0
            if asset_key in asset_assessment:
                asset_row = asset_assessment[asset_key]
                asset_values = []
                for i in [0, 1, 2, 3]:
                    val = asset_row.get(str(i), "")
                    if val and str(val) != "0":
                        try:
                            asset_values.append(float(val))
                        except (ValueError, TypeError):
                            pass
                
                if asset_values:  # Calcola solo se ci sono valori validi
                    asset_quadratic_mean = math.sqrt(sum(x**2 for x in asset_values) / len(asset_values))
                    asset_likelihood = (asset_quadratic_mean - 1) / 4
                    asset_likelihood = max(0.0, min(1.0, asset_likelihood))

            if asset_likelihood < 0:
                continue

            threat_likelihood_cat = self.value_to_category(threat_likelihood)
            asset_likelihood_cat = self.value_to_category(asset_likelihood)
            combined_likelihood_cat = self.RISK_MATRIX.get((threat_likelihood_cat, asset_likelihood_cat), threat_likelihood_cat)

            # --- Impact ---
            # Threat-specific impact (ultimi 2 criteri)
            threat_impact_values = []
            for i in [5, 6]:
                val = asset_data.get(str(i), "")
                if val and str(val) != "0":
                    try:
                        threat_impact_values.append(float(val))
                    except (ValueError, TypeError):
                        pass
            
            if not threat_impact_values:  # Se non ci sono valori validi, skip
                continue
                
            threat_impact_mean = math.sqrt(sum(x**2 for x in threat_impact_values) / len(threat_impact_values))
            threat_impact = (threat_impact_mean - 1) / 4
            threat_impact = max(0.0, min(1.0, threat_impact))

            # Asset impact (ultimi 5 criteri)
            asset_impact = -1.0
            if asset_key in asset_assessment:
                asset_row = asset_assessment[asset_key]
                asset_impact_values = []
                for i in [4, 5, 6, 7, 8]:
                    val = asset_row.get(str(i), "")
                    if val and str(val) != "0":
                        try:
                            asset_impact_values.append(float(val))
                        except (ValueError, TypeError):
                            pass
                
                if asset_impact_values:  # Calcola solo se ci sono valori validi
                    asset_impact_mean = math.sqrt(sum(x**2 for x in asset_impact_values) / len(asset_impact_values))
                    asset_impact = (asset_impact_mean - 1) / 4
                    asset_impact = max(0.0, min(1.0, asset_impact))

            if asset_impact < 0:
                continue

            threat_impact_cat = self.value_to_category(threat_impact)
            asset_impact_cat = self.value_to_category(asset_impact)
            combined_impact_cat = self.RISK_MATRIX.get((threat_impact_cat, asset_impact_cat), threat_impact_cat)

            # --- Risk ---
            risk_cat = self.RISK_MATRIX.get((combined_likelihood_cat, combined_impact_cat), "")

            priority = risk_priorities.get(risk_cat, 0)
            if priority > max_priority:
                max_priority = priority
                best_likelihood = combined_likelihood_cat
                best_impact = combined_impact_cat
                best_risk = risk_cat

        return best_likelihood, best_impact, best_risk

    def get_asset_likelihood_for_key(self, key):
        """Get asset likelihood for a specific asset key from the latest asset assessment"""
        if not self.asset_data:
            return -1.0
        
        # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
        assessment_keys = [k for k in self.asset_data.keys() if k.startswith('assessment_')]
        imported_keys = [k for k in self.asset_data.keys() if k.startswith('imported_')]
        
        # Use the latest assessment key if available, otherwise use latest imported key
        if assessment_keys:
            latest_key = max(assessment_keys)
        elif imported_keys:
            latest_key = max(imported_keys)
        else:
            latest_key = max(self.asset_data.keys()) if self.asset_data else None
        
        if not latest_key or latest_key not in self.asset_data:
            return -1.0
        
        asset_assessment = self.asset_data[latest_key]
        
        if key not in asset_assessment:
            return -1.0
        
        asset_data = asset_assessment[key]
        
        # Calculate asset likelihood from first 4 criteria (Dependency, Penetration, Cyber Maturity, Trust)
        try:
            values = []
            for i in [0, 1, 2, 3]:
                # Use get() with default to safely access the value
                val = asset_data.get(str(i), "")
                # Skip empty or zero values but continue with remaining criteria
                if val and str(val).strip() and str(val) != "0":
                    try:
                        values.append(float(val))
                    except (ValueError, TypeError):
                        continue  # Skip invalid values but continue processing
            
            # Require at least one valid value to calculate likelihood
            if not values:
                return -1.0
            
            # Use quadratic mean for asset likelihood
            quadratic_mean = math.sqrt(sum(x**2 for x in values) / len(values))
            likelihood = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            
            return max(0.0, min(1.0, likelihood))
            
        except Exception as e:
            # Catch any unexpected errors and return -1.0
            return -1.0

    def get_asset_impact_for_key(self, key):
        """Get asset impact for a specific asset key from the latest asset assessment"""
        if not self.asset_data:
            return -1.0
        
        # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
        assessment_keys = [k for k in self.asset_data.keys() if k.startswith('assessment_')]
        imported_keys = [k for k in self.asset_data.keys() if k.startswith('imported_')]
        
        # Use the latest assessment key if available, otherwise use latest imported key
        if assessment_keys:
            latest_key = max(assessment_keys)
        elif imported_keys:
            latest_key = max(imported_keys)
        else:
            latest_key = max(self.asset_data.keys()) if self.asset_data else None
        
        if not latest_key or latest_key not in self.asset_data:
            return -1.0
        
        asset_assessment = self.asset_data[latest_key]
        
        if key not in asset_assessment:
            return -1.0
        
        asset_data = asset_assessment[key]
        
        # Calculate asset impact from last 5 criteria (Performance, Schedule, Costs, Reputation, Recovery)
        try:
            values = []
            for i in [4, 5, 6, 7, 8]:
                # Use get() with default to safely access the value
                val = asset_data.get(str(i), "")
                # Skip empty or zero values but continue with remaining criteria
                if val and str(val).strip() and str(val) != "0":
                    try:
                        values.append(float(val))
                    except (ValueError, TypeError):
                        continue  # Skip invalid values but continue processing
            
            # Require at least one valid value to calculate impact
            if not values:
                return -1.0
            
            # Use quadratic mean for asset impact
            quadratic_mean = math.sqrt(sum(x**2 for x in values) / len(values))
            impact = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
            
            return max(0.0, min(1.0, impact))
            
        except Exception as e:
            # Catch any unexpected errors and return -1.0
            return -1.0

    def setup_combobox_styles(self):
        """Configure custom styles for Comboboxes with criteria colors"""
        style = ttk.Style()
        
        # Use a compatible theme
        available_themes = style.theme_names()
        if 'clam' in available_themes:
            style.theme_use('clam')
        elif 'alt' in available_themes:
            style.theme_use('alt')
        
        # Configure a style for each criterion with its specific color
        for i, color in enumerate(self.CRITERIA_COLORS):
            style_name = f"Criteria{i}.TCombobox"
            try:
                # Configure the style for the Combobox
                style.configure(style_name,
                               fieldbackground=color,
                               background=color,
                               selectbackground=color,
                               foreground='black',
                               selectforeground='black',
                               insertcolor='black',
                               bordercolor=self.COLORS['gray'],
                               arrowcolor=self.COLORS['dark'],
                               focuscolor='none',
                               relief='flat',
                               borderwidth=1)
                
                # Configure states to maintain color and black text
                style.map(style_name,
                         fieldbackground=[('readonly', color),
                                        ('disabled', color),
                                        ('active', color),
                                        ('focus', color),
                                        ('!focus', color)],
                         background=[('readonly', color),
                                   ('disabled', color),
                                   ('active', color),
                                   ('pressed', color)],
                         selectbackground=[('readonly', color),
                                         ('disabled', color),
                                         ('active', color),
                                         ('focus', color)],
                         foreground=[('readonly', 'black'),
                                   ('disabled', 'black'),
                                   ('active', 'black'),
                                   ('focus', 'black'),
                                   ('!focus', 'black')],
                         selectforeground=[('readonly', 'black'),
                                         ('disabled', 'black'),
                                         ('active', 'black'),
                                         ('focus', 'black'),
                                         ('selected', 'black')])
                
            except Exception as e:
                #print(f"Error configuring style {style_name}: {e}")
                # Fallback: configure only basic properties                    
                style.configure(style_name, 
                                fieldbackground=color, 
                                foreground='black',
                                selectforeground='black')

    def load_threat_details(self):
        """Load threat details from Threat_Subset.csv"""
        threat_details = {}
        threats_file = os.path.join(get_base_path(), "Threat_Subset.csv")
        
        try:
            with open(threats_file, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                for row in reader:
                    threat_name = row.get('THREAT', '').strip()
                    if threat_name:
                        threat_details[threat_name] = {
                            'category': row.get('THREAT CATEGORY', '').strip(),
                            'description': row.get('THREAT DESCRIPTION', '').strip()
                        }
        except FileNotFoundError:
            logging.warning(f"Threat details file not found: {threats_file}")
        except Exception as e:
            logging.error(f"Error loading threat details: {e}")
        
        return threat_details

    def get_analyzed_threats(self):
        """Get list of threats that have been analyzed (have at least one non-empty risk value)"""
        analyzed_threats = []
        
        for threat_name in self.threat_data.keys():
            # Check if this threat has at least one valid risk calculation
            threat_data = self.threat_data[threat_name]
            has_valid_risk = False
            
            for asset_key, asset_data in threat_data.items():
                # Calculate likelihood and impact
                likelihood = self.calculate_likelihood_from_saved_data(asset_data)
                impact = self.calculate_impact_from_saved_data(asset_data)
                
                # If both are valid, we have a risk value
                if likelihood >= 0 and impact >= 0:
                    has_valid_risk = True
                    break
            
            if has_valid_risk:
                analyzed_threats.append(threat_name)
        
        return analyzed_threats

    def get_analyzed_assets(self):
        """Get list of assets that have been analyzed (either through threats or asset assessment)"""
        analyzed_assets = set()
        
        # Get assets from threat analysis
        for threat_name in self.threat_data.keys():
            threat_data = self.threat_data[threat_name]
            
            for asset_key, asset_data in threat_data.items():
                # Check if this combination has valid data
                likelihood = self.calculate_likelihood_from_saved_data(asset_data)
                impact = self.calculate_impact_from_saved_data(asset_data)
                
                if likelihood >= 0 and impact >= 0:
                    # Extract asset name from asset_key (format: "1_probability" -> asset index 0)
                    asset_index = int(asset_key.split('_')[0]) - 1
                    if 0 <= asset_index < len(self.ASSET_CATEGORIES):
                        asset_name = self.ASSET_CATEGORIES[asset_index][2]  # Component name
                        analyzed_assets.add(asset_name)
        
        # Also get assets from asset assessment (independent of threat analysis)
        if self.asset_data:
            # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
            assessment_keys = [key for key in self.asset_data.keys() if key.startswith('assessment_')]
            imported_keys = [key for key in self.asset_data.keys() if key.startswith('imported_')]
            
            # Use the latest assessment key if available, otherwise use latest imported key
            if assessment_keys:
                latest_key = max(assessment_keys)
            elif imported_keys:
                latest_key = max(imported_keys)
            else:
                latest_key = max(self.asset_data.keys())
                
            if latest_key in self.asset_data:
                asset_assessment = self.asset_data[latest_key]
                
                for asset_key, asset_data in asset_assessment.items():
                    if asset_key.endswith('_probability'):
                        # Extract asset index and name
                        asset_index = int(asset_key.split('_')[0]) - 1
                        if 0 <= asset_index < len(self.ASSET_CATEGORIES):
                            asset_name = self.ASSET_CATEGORIES[asset_index][2]
                            
                            # Check if this asset has likelihood and impact
                            likelihood_cat, impact_cat = self.get_asset_likelihood_impact(asset_name)
                            if likelihood_cat and impact_cat:
                                analyzed_assets.add(asset_name)
        
        return list(analyzed_assets)

    def get_threat_max_risk(self, threat_name):
        """Get maximum risk values for a threat (same logic as main table update)"""
        risk_priorities = {"Very High": 5, "High": 4, "Medium": 3, "Low": 2, "Very Low": 1, "": 0}
        
        max_likelihood = ""
        max_impact = ""
        max_risk = ""
        max_priority = 0
        
        if threat_name not in self.threat_data:
            return max_likelihood, max_impact, max_risk
        
        threat_data = self.threat_data[threat_name]
        
        for asset_key, asset_data in threat_data.items():
            # Calculate likelihood and impact
            likelihood = self.calculate_likelihood_from_saved_data(asset_data)
            impact = self.calculate_impact_from_saved_data(asset_data)
            
            # Calculate risk if both are available
            if likelihood >= 0 and impact >= 0:
                likelihood_cat = self.value_to_category(likelihood)
                impact_cat = self.value_to_category(impact)
                risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
                
                priority = risk_priorities.get(risk_level, 0)
                if priority > max_priority:
                    max_priority = priority
                    max_likelihood = likelihood_cat
                    max_impact = impact_cat
                    max_risk = risk_level
        
        return max_likelihood, max_impact, max_risk

    def get_threat_asset_risk(self, threat_name, asset_name):
        """Get risk values for a specific threat-asset combination"""
        if threat_name not in self.threat_data:
            return "", "", ""
        
        # Find asset index by name
        asset_index = -1
        for i, (category, sub_category, component) in enumerate(self.ASSET_CATEGORIES):
            if component == asset_name:
                asset_index = i
                break
        
        if asset_index == -1:
            return "", "", ""
        
        asset_key = f"{asset_index + 1}_probability"
        threat_data = self.threat_data[threat_name]
        
        if asset_key not in threat_data:
            return "", "", ""
        
        asset_data = threat_data[asset_key]
        
        # Calculate likelihood and impact
        likelihood = self.calculate_likelihood_from_saved_data(asset_data)
        impact = self.calculate_impact_from_saved_data(asset_data)
        
        if likelihood >= 0 and impact >= 0:
            likelihood_cat = self.value_to_category(likelihood)
            impact_cat = self.value_to_category(impact)
            risk_level = self.RISK_MATRIX.get((likelihood_cat, impact_cat), "")
            return likelihood_cat, impact_cat, risk_level
        
        return "", "", ""

    def get_asset_likelihood_impact(self, asset_name):
        """Get asset likelihood and impact values from the latest asset assessment"""
        if not self.asset_data:
            return "", ""
        
        # Find the most recent assessment - prioritize assessment_ keys over imported_ keys
        assessment_keys = [key for key in self.asset_data.keys() if key.startswith('assessment_')]
        imported_keys = [key for key in self.asset_data.keys() if key.startswith('imported_')]
        
        # Use the latest assessment key if available, otherwise use latest imported key
        if assessment_keys:
            latest_key = max(assessment_keys)
        elif imported_keys:
            latest_key = max(imported_keys)
        else:
            latest_key = max(self.asset_data.keys()) if self.asset_data else None
        
        if not latest_key or latest_key not in self.asset_data:
            return "", ""
        
        # Find asset index by name
        asset_index = -1
        for i, (category, sub_category, component) in enumerate(self.ASSET_CATEGORIES):
            if component == asset_name:
                asset_index = i
                break
        
        if asset_index == -1:
            return "", ""
        
        asset_key = f"{asset_index + 1}_probability"
        asset_assessment = self.asset_data[latest_key]
        
        if asset_key not in asset_assessment:
            return "", ""
        
        asset_data = asset_assessment[asset_key]
        
        try:
            # Calculate likelihood from first 4 criteria (Dependency, Penetration, Cyber Maturity, Trust)
            likelihood_values = []
            for i in [0, 1, 2, 3]:
                val = asset_data.get(str(i), "")
                if val and val != "0":
                    likelihood_values.append(float(val))
            
            # Calculate impact from last 5 criteria (Performance, Schedule, Costs, Reputation, Recovery)
            impact_values = []
            for i in [4, 5, 6, 7, 8]:
                val = asset_data.get(str(i), "")
                if val and val != "0":
                    impact_values.append(float(val))
            
            # Calculate likelihood if we have all 4 values
            likelihood_cat = ""
            if len(likelihood_values) > 0:
                quadratic_mean = math.sqrt(sum(x**2 for x in likelihood_values) / len(likelihood_values))
                likelihood = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
                likelihood = max(0.0, min(1.0, likelihood))
                likelihood_cat = self.value_to_category(likelihood)
            
            # Calculate impact if we have all 5 values
            impact_cat = ""
            if len(impact_values) > 0:
                quadratic_mean = math.sqrt(sum(x**2 for x in impact_values) / len(impact_values))
                impact = (quadratic_mean - 1) / 4  # Normalize [1,5] -> [0,1]
                impact = max(0.0, min(1.0, impact))
                impact_cat = self.value_to_category(impact)
            
            return likelihood_cat, impact_cat
            
        except (ValueError, KeyError):
            return "", ""

    def load_controls_for_threat(self, threat_name):
        """Load controls from Control.csv that address the specified threat"""
        controls = []
        controls_file = os.path.join(get_base_path(), "Control.csv")
        
        try:
            with open(controls_file, 'r', newline='', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile, delimiter=';')
                for row in reader:
                    try:
                        # Check if threat is addressed by this control (column "Threats addressed")
                        threats_addressed = row.get('Threats addressed', '')
                        if threats_addressed and isinstance(threats_addressed, str):
                            threats_addressed = threats_addressed.strip()
                            if threats_addressed:
                                # Split by comma and clean each threat name
                                threat_names = [t.strip() for t in threats_addressed.split(',')]
                                
                                # Check for exact match (case-insensitive) with any of the threats in the list
                                threat_found = False
                                for addressed_threat in threat_names:
                                    if isinstance(addressed_threat, str):
                                        if addressed_threat.lower() == threat_name.lower():
                                            threat_found = True
                                            break
                                        # Also check if our threat name is contained in the addressed threat
                                        # (for cases like "Malicious code/software/activity: Network exploit")
                                        elif threat_name.lower() in addressed_threat.lower():
                                            threat_found = True
                                            break
                                
                                if threat_found:
                                    controls.append({
                                        'title': str(row.get('Control title', '')).strip(),
                                        'control': str(row.get('Control', '')).strip(),
                                        'description': str(row.get('Control description', '')).strip(),
                                        'reference': str(row.get('Reference frameworks', '')).strip(),
                                        'lifecycle': str(row.get('Lifecycle phase', '')).strip(),
                                        'segment': str(row.get('Segment', '')).strip(),
                                        'criterio': str(row.get('Criterio', '')).strip(),
                                    })
                    except Exception as e:
                        print(f"Error processing control row: {e}")
                        continue
                            
        except FileNotFoundError:
            print(f"Controls file not found: {controls_file}")
            logging.warning(f"Controls file not found: {controls_file}")
        except Exception as e:
            print(f"Error loading controls: {e}")
            logging.error(f"Error loading controls: {e}")
        
        return controls

    def get_controls_status_for_threat(self, threat_name):
        """Get controls status for a threat: applied vs available"""
        try:
            print(f"[DEBUG] Getting controls status for threat: {threat_name}")
            print(f"[DEBUG] Applied controls list: {getattr(self, 'applied_controls', [])}")
            
            # Get all available controls for this threat from CSV
            available_controls = self.load_controls_for_threat(threat_name)
            print(f"[DEBUG] Available controls from CSV: {len(available_controls)}")
            
            # Get applied controls from the applied_controls list (persistent data)
            applied_controls = []
            available_not_applied = []
            
            # Use the persistent applied_controls list instead of selected_controls
            if hasattr(self, 'applied_controls') and self.applied_controls:
                print(f"[DEBUG] Processing {len(self.applied_controls)} applied controls")
                for control_id in self.applied_controls:
                    print(f"[DEBUG] Looking for control ID: {control_id}")
                    # Find this control in our CONTROLS list
                    if hasattr(self, 'CONTROLS') and self.CONTROLS:
                        control_data = next((c for c in self.CONTROLS if c['id'] == control_id), None)
                        if control_data:
                            print(f"[DEBUG] Found control: {control_data.get('title', 'N/A')}")
                            print(f"[DEBUG] Control reference: '{control_data.get('reference', '')}'")
                            print(f"[DEBUG] Control lifecycle: '{control_data.get('lifecycle', '')}'")
                            print(f"[DEBUG] All control keys: {list(control_data.keys())}")
                            # Check if this control addresses the current threat
                            threats_addressed = control_data.get('threats_addressed', '')
                            print(f"[DEBUG] Control addresses threats: {threats_addressed}")
                            if threats_addressed and isinstance(threats_addressed, str):
                                threats_list = [t.strip().lower() for t in threats_addressed.split(',')]
                                matches_threat = any(threat_name.lower() in threat.lower() or threat.lower() in threat_name.lower() 
                                      for threat in threats_list)
                                print(f"[DEBUG] Threat matches: {matches_threat}")
                                if matches_threat:
                                    applied_controls.append({
                                        'id': control_data.get('id', 'N/A'),
                                        'title': control_data.get('title', ''),
                                        'description': control_data.get('description', ''),
                                        'criteria': control_data.get('criteria', ''),
                                        'cluster': control_data.get('cluster', ''),
                                        'reference': control_data.get('reference', ''),
                                        'lifecycle': control_data.get('lifecycle', ''),
                                        'segment': control_data.get('segment', '')
                                    })
                                    print(f"[DEBUG] Added to applied controls")
                        else:
                            print(f"[DEBUG] Control ID {control_id} not found in CONTROLS dataset")
            else:
                print(f"[DEBUG] No applied controls found")
            
            print(f"[DEBUG] Final applied controls count: {len(applied_controls)}")
            
            # For available controls, mark which ones are not applied
            applied_ids = {c['id'] for c in applied_controls}
            
            for i, control in enumerate(available_controls):
                try:
                    # Try to find the corresponding control in CONTROLS to get the ID
                    matching_control = None
                    if hasattr(self, 'CONTROLS') and self.CONTROLS:
                        for c in self.CONTROLS:
                            if (c.get('title') == control.get('title') or 
                                c.get('description') == control.get('description')):
                                matching_control = c
                                break
                    
                    control_id = matching_control.get('id', 'N/A') if matching_control else 'N/A'
                    control_cluster = matching_control.get('cluster', 'N/A') if matching_control else 'N/A'
                    
                    if control_id not in applied_ids:
                        available_not_applied.append({
                            'id': control_id,
                            'title': control.get('title', ''),
                            'description': control.get('control', ''),  # Note: this maps to 'description' 
                            'criteria': control.get('criterio', ''),
                            'cluster': control_cluster,  # Add cluster information
                            'reference': control.get('reference', ''),
                            'lifecycle': control.get('lifecycle', ''),
                            'segment': control.get('segment', '')
                        })
                except Exception as e:
                    continue
            
            result = {
                'applied': applied_controls,
                'available_not_applied': available_not_applied,
                'total_available': len(available_controls)
            }
            return result
        
        except Exception as e:
            import traceback
            traceback.print_exc()
            return {
                'applied': [],
                'available_not_applied': [],
                'total_available': 0
            }


def main():
    """Main function to run the application"""
    root = tk.Tk()
    app = RiskAssessmentTool(root)
    root.mainloop()


if __name__ == "__main__":
    main()
