"""
Attack Graph Analyzer for Space Systems Cybersecurity
Attack graph analyzer for space systems cybersecurity

This script converts the CSV of threat relationships into a NetworkX graph
and provides analysis and visualization functionality.

🔧 QUICK CONFIGURATION:
When you run this script, a file dialog will appear allowing you to select
the CSV file with threats to analyze. Alternatively, you can still modify the 
THREAT_FILE_NAME variable in the "ANALYSIS CONFIGURATION" section (around line 58).

📋 THREAT FILE FORMAT:
The threat CSV file must contain the following columns (separated by ';'):
- THREAT: Threat name
- Likelihood: Threat likelihood (Very Low, Low, Medium, High, Very High)
- Impact: Threat impact (Very Low, Low, Medium, High, Very High)  
- Risk: Threat risk (Very Low, Low, Medium, High, Very High)

Example of supported file:
THREAT;Likelihood;Impact;Risk
Social Engineering;High;Very High;Very High
Unauthorized physical access;Medium;Very High;High
Seizure of control: Satellite bus;Low;Very High;Medium
"""

# Risk Assessment Tool - Relation between Threats
# Purpose: Analyze the relationships between threats in space systems and create a threat attack graph
# Author: Thesis work for space program risk assessment tool Giuseppe Nonni 1948023 giuseppe.nonni@gmail.com

import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from matplotlib.lines import Line2D
from matplotlib.patches import FancyBboxPatch
import numpy as np
from collections import Counter
import warnings
import os
import sys
import traceback
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
warnings.filterwarnings('ignore')

def get_base_path():
    """Get the base path for the application (works with both .py and .exe)"""
    if getattr(sys, 'frozen', False):
        # Running as compiled executable
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

# Conditional import for scipy
try:
    import scipy
    HAS_SCIPY = True
except ImportError:
    HAS_SCIPY = False

# =============================================================================
# GUI STYLING CONFIGURATION - Consistent with other components
# =============================================================================

COLORS = {
    'primary': '#4a90c2', 'success': '#28a745', 'white': '#ffffff',
    'light': '#f8f9fa', 'dark': '#2c3e50', 'gray': '#6c757d',
    'criteria_header': '#5a67d8', 'criteria_bg': '#edf2f7',
    'warning': '#ffc107', 'danger': '#dc3545', 'info': '#17a2b8'
}

# =============================================================================
# ANALYSIS CONFIGURATION - MODIFY THESE VALUES TO CUSTOMIZE THE ANALYSIS
# =============================================================================

# 🔧 MAIN PARAMETER: CSV file name with threats to analyze
# This variable can be modified to change the threat source file programmatically
# OR you can select the file interactively when running the script
# IMPORTANT: Only threats that are present BOTH in relations AND in this file will be analyzed
THREAT_FILE_NAME = "CSV_Export_20250708_094356/Threat_Analyzed.csv"  # Simplified CSV file: THREAT;Likelihood;Impact;Risk

# These will be dynamically calculated after loading the graph
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": None,  # Will be set to the threat with most outgoing connections
    "target_threat": None,  # Will be set to the threat with most incoming connections
    "max_path_length": 5
}

# Flag to decide whether to save the plot of all paths (1) or only the combined one (0)
save_path = 0
# Flag to decide whether to save maximum 5 combined paths (1) or all combined paths (0) 
max_five = 0
SPECIFIC_THREAT = None  # Will be set to the threat with highest risk

# This will be dynamically populated with the 6 most critical paths
MULTIPLE_PATH_ANALYSIS = []

# Analysis parameters
ANALYSIS_PARAMETERS = {
    "max_paths_per_pair": 3,
    "max_critical_path_length": 6,
    "top_centrality_nodes": 5,
    "top_critical_paths": 10
}

# Configuration for analyzing connections of a specific threat
THREAT_CONNECTION_ANALYSIS = {
    "target_threat": SPECIFIC_THREAT,  # Change this to analyze a different threat
    "max_distance": 1,  # Maximum distance: 1=direct connections, 2=two-level connections
    "show_relation_types": True,  # Show relation types
    "include_predecessors": True,  # Analyze threats that lead to the target
    "include_successors": True,   # Analyze threats enabled by the target
    "save_visualization": True   # Save a connections graph
}

# Configuration for star graph - shows all nodes connected to a specific threat
STAR_GRAPH_CONFIG = {
    "center_threat": SPECIFIC_THREAT,  # Change this to analyze a different threat
    "enable_star_graph": True,  # Set to False to disable
    "max_distance": 1,  # Maximum distance from central node (1=direct neighbors, 2=neighbors of neighbors)
    "show_edge_labels": True  # Show labels on connections
}

class OutputManager:
    """Manages output to text file."""
    
    def __init__(self, output_file="attack_graph_analysis.txt"):
        # Create Output directory if it doesn't exist
        output_dir = os.path.join(get_output_path(), "Output")
        os.makedirs(output_dir, exist_ok=True)
        
        self.output_file = os.path.join(output_dir, output_file)
        self.file_handle = None
        self.start_logging()
    
    def start_logging(self):
        """Starts output logging."""
        try:
            self.file_handle = open(self.output_file, 'w', encoding='utf-8')
            self.log(f"=== ATTACK GRAPH ANALYSIS - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===\n")
        except Exception as e:
            self.file_handle = None
    
    def log(self, message):
        """Writes a message both to console and file."""
        if self.file_handle:
            try:
                self.file_handle.write(message + '\n')
                self.file_handle.flush()
            except Exception:
                pass
    
    def close(self):
        """Closes the output file."""
        if self.file_handle:
            try:
                self.file_handle.write(f"\n=== ANALYSIS COMPLETED - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')} ===")
                self.file_handle.close()
                ##print(f"\n📄 Report saved to: {self.output_file}")
            except Exception:
                pass


class AttackGraphAnalyzer:    
    def __init__(self, csv_file_path, subset_file_path="Threat_Analyzed.csv", output_file="attack_graph_analysis.txt"):
        """
        Initializes the attack graph analyzer.
        
        Args:
            csv_file_path (str): Path to the CSV file with threat relationships
            subset_file_path (str): Path to the CSV file with the subset of threats to analyze
            output_file (str): Name of the output file for the report
        """
        self.csv_file_path = csv_file_path
        self.subset_file_path = subset_file_path
        self.df = None
        self.subset_threats = None
        self.graph = None
        
        self.output = OutputManager(output_file)
        self.load_data()
        self.load_subset()
        self.create_graph()
        
        # Calculate dynamic configurations after the graph is created
        self._calculate_dynamic_configurations()

    def load_data(self):
        """Loads data from CSV file."""
        try:
            self.df = pd.read_csv(self.csv_file_path, sep=';')
            self.output.log(f"Data loaded successfully: {len(self.df)} relationships found")
            self.output.log(f"Columns: {list(self.df.columns)}")
        except Exception as e:
            self.output.log(f"Error loading file: {e}")
            return
    
    def load_subset(self):
        """Loads the subset of threats to analyze from the THREAT_FILE_NAME file."""
        try:
            if os.path.exists(self.subset_file_path):
                subset_df = pd.read_csv(self.subset_file_path, sep=';')
                
                # Check that the THREAT column exists
                if 'THREAT' not in subset_df.columns:
                    self.output.log(f"❌ Error: the file {THREAT_FILE_NAME} must contain a 'THREAT' column")
                    self.output.log(f"   Columns found: {list(subset_df.columns)}")
                    self.subset_threats = None
                    return
                
                self.subset_threats = set(subset_df['THREAT'].tolist())
                self.output.log(f"✅ File {THREAT_FILE_NAME} loaded successfully")
                self.output.log(f"📋 Subset loaded: {len(self.subset_threats)} threats selected")
                self.output.log(f"🎯 Only threats present BOTH in relationships AND in {THREAT_FILE_NAME} will be analyzed")
                
                # Show the complete list of loaded threats (sorted)
                threat_list = sorted(list(self.subset_threats))
                self.output.log(f"📝 Threats loaded from {THREAT_FILE_NAME}:")
                for i, threat in enumerate(threat_list, 1):
                    self.output.log(f"   {i:2d}. {threat}")
                
                if len(self.subset_threats) == 0:
                    self.output.log(f"⚠️  The file {THREAT_FILE_NAME} is empty or does not contain valid threats")
                    
            else:
                self.output.log(f"⚠️  Subset file '{self.subset_file_path}' not found. Analysis on all relationship threats.")
                self.subset_threats = None
        except Exception as e:
            self.output.log(f"❌ Error loading subset: {e}")
            self.output.log(f"   Details: {str(e)}")
            self.subset_threats = None
    
    def _is_threat_in_subset(self, threat_name):
        """
        Checks if a threat is present in the THREAT_FILE_NAME file.
        Only threats that are present BOTH in relationships AND in the subset file will be included in the analysis.
        """
        if self.subset_threats is None:
            return True  # If not in the subset include all the threat
        return threat_name in self.subset_threats
    
    def _filter_graph_by_subset(self):
        """Filters the graph to include only threats present BOTH in relationships AND in the THREAT_FILE_NAME file."""
        if self.subset_threats is None or self.graph is None:
            self.output.log("⚠️  No filter applied: subset file not available or empty graph")
            return  # No filter to apply
        
        # Identify nodes to remove: keep only those that are in the subset AND in relationships
        nodes_to_remove = []
        nodes_in_graph = set(self.graph.nodes())
        
        for node in nodes_in_graph:
            # Remove the node if it is NOT present in the THREAT_FILE_NAME file
            if not self._is_threat_in_subset(node):
                nodes_to_remove.append(node)
        
        # Remove nodes not in subset
        self.graph.remove_nodes_from(nodes_to_remove)
        
        self.output.log(f"📊 FILTER APPLIED:")
        self.output.log(f"   Threats in {THREAT_FILE_NAME} file: {len(self.subset_threats)}")
        self.output.log(f"   Threats in relationships (original): {len(nodes_in_graph)}")
        self.output.log(f"   Threats removed (not in {THREAT_FILE_NAME}): {len(nodes_to_remove)}")
        self.output.log(f"   Final threats (intersection): {self.graph.number_of_nodes()}")
        self.output.log(f"   Final relationships: {self.graph.number_of_edges()}")
        
        if len(nodes_to_remove) > 0:
            self.output.log(f"🗑️  Threats removed: {sorted(nodes_to_remove)[:10]}{'...' if len(nodes_to_remove) > 10 else ''}")

    def create_graph(self):
        """Creates the directed NetworkX graph from the DataFrame."""
        if self.df is None:
            self.output.log("No data available to create the graph")
            return
        
        # Create a directed graph
        self.graph = nx.DiGraph()
        
        # Add nodes and edges
        for _, row in self.df.iterrows():
            source = row['Source Threat']
            target = row['Target Threat']
            source_cat = row['Source Category']
            target_cat = row['Target Category']
            relation_type = row['Relation Type']
            
            # Add nodes with attributes
            self.graph.add_node(source, category=source_cat)
            self.graph.add_node(target, category=target_cat)
            
            # Add edge with attributes
            self.graph.add_edge(source, target, 
                              relation_type=relation_type,
                              source_category=source_cat,
                              target_category=target_cat)
        
        self.output.log(f"Graph created with {self.graph.number_of_nodes()} nodes and {self.graph.number_of_edges()} edges")
        
        # Apply the subset filter
        self._filter_graph_by_subset()
    
    def _calculate_dynamic_configurations(self):
        """
        Calculate dynamic configurations based on the loaded graph.
        This function adapts analysis parameters based on graph characteristics.
        """
        if self.graph is None:
            self.output.log("⚠️ Cannot calculate dynamic configurations: graph not available")
            return
        
        # Get basic graph metrics
        num_nodes = self.graph.number_of_nodes()
        num_edges = self.graph.number_of_edges()
        
        self.output.log(f"\n🔧 CALCULATING DYNAMIC CONFIGURATIONS")
        self.output.log(f"   Graph size: {num_nodes} nodes, {num_edges} edges")
        
        # Update global ANALYSIS_PARAMETERS based on graph size
        global ANALYSIS_PARAMETERS
        
        # Adjust parameters based on graph size
        if num_nodes < 50:
            # Small graph - more detailed analysis
            ANALYSIS_PARAMETERS["top_centrality_nodes"] = min(10, max(5, num_nodes // 2))
            ANALYSIS_PARAMETERS["max_paths_per_pair"] = 5
            ANALYSIS_PARAMETERS["max_critical_path_length"] = 6
            ANALYSIS_PARAMETERS["top_critical_paths"] = min(15, num_nodes)
            
        elif num_nodes < 200:
            # Medium graph - balanced analysis
            ANALYSIS_PARAMETERS["top_centrality_nodes"] = min(15, num_nodes // 4)
            ANALYSIS_PARAMETERS["max_paths_per_pair"] = 3
            ANALYSIS_PARAMETERS["max_critical_path_length"] = 5
            ANALYSIS_PARAMETERS["top_critical_paths"] = min(20, num_nodes // 2)
            
        else:
            # Large graph - focus on most important elements
            ANALYSIS_PARAMETERS["top_centrality_nodes"] = min(20, num_nodes // 8)
            ANALYSIS_PARAMETERS["max_paths_per_pair"] = 2
            ANALYSIS_PARAMETERS["max_critical_path_length"] = 4
            ANALYSIS_PARAMETERS["top_critical_paths"] = min(25, num_nodes // 4)
        
        # Dynamic threat selection based on available threats
        available_threats = list(self.graph.nodes())
        
        # Update SPECIFIC_PATH_ANALYSIS with dynamic threat selection
        global SPECIFIC_PATH_ANALYSIS
        
        # Find good source threats (high out-degree, low in-degree)
        out_degrees = dict(self.graph.out_degree())
        in_degrees = dict(self.graph.in_degree())
        
        # Potential sources: high out-degree, low in-degree
        source_candidates = [(node, out_degrees[node], in_degrees[node]) 
                           for node in available_threats 
                           if out_degrees[node] > 0]
        source_candidates.sort(key=lambda x: (x[1], -x[2]), reverse=True)
        
        # Potential targets: high in-degree, low out-degree  
        target_candidates = [(node, in_degrees[node], out_degrees[node])
                           for node in available_threats
                           if in_degrees[node] > 0]
        target_candidates.sort(key=lambda x: (x[1], -x[2]), reverse=True)
        
        # Update source and target if we found good candidates
        if source_candidates:
            best_source = source_candidates[0][0]
            SPECIFIC_PATH_ANALYSIS["source_threat"] = best_source
            self.output.log(f"   📍 Dynamic source selected: {best_source}")
            
        if target_candidates:
            best_target = target_candidates[0][0]
            SPECIFIC_PATH_ANALYSIS["target_threat"] = best_target
            self.output.log(f"   🎯 Dynamic target selected: {best_target}")
        
        # Adjust path length based on graph density
        density = nx.density(self.graph)
        if density > 0.3:  # High density
            SPECIFIC_PATH_ANALYSIS["max_path_length"] = 3
        elif density > 0.1:  # Medium density
            SPECIFIC_PATH_ANALYSIS["max_path_length"] = 4
        else:  # Low density
            SPECIFIC_PATH_ANALYSIS["max_path_length"] = 5
            
        # Update STAR_GRAPH_CONFIG with dynamic center threat selection
        global STAR_GRAPH_CONFIG
        
        # Find threat with highest betweenness centrality as center
        if num_nodes > 2:  # Need at least 3 nodes for meaningful centrality
            try:
                betweenness_centrality = nx.betweenness_centrality(self.graph)
                if betweenness_centrality:
                    center_threat = max(betweenness_centrality.keys(), 
                                      key=lambda x: betweenness_centrality[x])
                    STAR_GRAPH_CONFIG["center_threat"] = center_threat
                    self.output.log(f"   ⭐ Dynamic center threat selected: {center_threat}")
            except Exception as e:
                self.output.log(f"   ⚠️ Could not calculate dynamic center threat: {e}")
        
        # Update MULTIPLE_PATH_ANALYSIS with dynamic paths
        global MULTIPLE_PATH_ANALYSIS
        MULTIPLE_PATH_ANALYSIS.clear()
        
        # Add multiple interesting paths based on graph analysis
        if len(source_candidates) >= 2 and len(target_candidates) >= 2:
            # Add top 3 source-target combinations
            for i in range(min(3, len(source_candidates))):
                for j in range(min(2, len(target_candidates))):
                    if i * 2 + j < 5:  # Limit total paths
                        source = source_candidates[i][0]
                        target = target_candidates[j][0]
                        if source != target:
                            MULTIPLE_PATH_ANALYSIS.append({
                                "description": f"High-criticality path #{i+1}-{j+1}",
                                "source": source,
                                "target": target
                            })
        
        # Log final configuration
        self.output.log(f"   ✅ Dynamic configuration completed:")
        self.output.log(f"      - Top centrality nodes: {ANALYSIS_PARAMETERS['top_centrality_nodes']}")
        self.output.log(f"      - Max paths per pair: {ANALYSIS_PARAMETERS['max_paths_per_pair']}")
        self.output.log(f"      - Max path length: {ANALYSIS_PARAMETERS['max_critical_path_length']}")
        self.output.log(f"      - Top critical paths: {ANALYSIS_PARAMETERS['top_critical_paths']}")
        self.output.log(f"      - Multiple paths configured: {len(MULTIPLE_PATH_ANALYSIS)}")

    def get_graph_statistics(self):
        """Calculates and displays graph statistics."""
        if self.graph is None:
            self.output.log("Graph not available")
            return {}
        
        stats = {
            'Number of nodes': self.graph.number_of_nodes(),
            'Number of edges': self.graph.number_of_edges(),
            'Graph density': nx.density(self.graph),
            'Is connected (weakly)': nx.is_weakly_connected(self.graph),
            'Is acyclic (DAG)': nx.is_directed_acyclic_graph(self.graph),
            'Number of connected components': nx.number_weakly_connected_components(self.graph)
        }
        
        self.output.log("\n=== GRAPH STATISTICS ===")
        for key, value in stats.items():
            self.output.log(f"{key}: {value}")
        
        # Degree statistics
        in_degrees = dict(self.graph.in_degree())
        out_degrees = dict(self.graph.out_degree())
        
        self.output.log(f"\nAverage in-degree: {np.mean(list(in_degrees.values())):.2f}")
        self.output.log(f"Average out-degree: {np.mean(list(out_degrees.values())):.2f}")
        
        # Top 5 nodes by in-degree (most common targets)
        top_targets = sorted(in_degrees.items(), key=lambda x: x[1], reverse=True)[:5]
        self.output.log("\n=== TOP 5 MOST TARGETED THREATS ===")
        for threat, degree in top_targets:
            self.output.log(f"{threat}: {degree} incoming attacks")
        
        # Top 5 nodes by out-degree (most common sources)
        top_sources = sorted(out_degrees.items(), key=lambda x: x[1], reverse=True)[:5]
        self.output.log("\n=== TOP 5 THREATS THAT ENABLE OTHERS ===")
        for threat, degree in top_sources:
            self.output.log(f"{threat}: {degree} outgoing attacks")
        
        return stats
    
    def analyze_categories(self):
        """Analyzes threat categories and their relationships."""
        if self.df is None:
            return
        
        self.output.log("\n=== CATEGORY ANALYSIS ===")
        
        # Count categories
        all_categories = list(self.df['Source Category']) + list(self.df['Target Category'])
        category_counts = Counter(all_categories)
        
        self.output.log("Category distribution:")
        for cat, count in category_counts.most_common():
            self.output.log(f"  {cat}: {count} occurrences")
        
        # Analyze relationship types
        relation_counts = Counter(self.df['Relation Type'])
        self.output.log("\nRelationship types:")
        for rel_type, count in relation_counts.most_common():
            self.output.log(f"  {rel_type}: {count} relationships")
        
        # Category relationship matrix
        category_relations = self.df.groupby(['Source Category', 'Target Category']).size().reset_index(name='count')
        self.output.log("\nRelationships between categories:")
        for _, row in category_relations.iterrows():
            self.output.log(f"  {row['Source Category']} → {row['Target Category']}: {row['count']} relationships")
    
    def analyze_centrality(self):
        """
        Analyzes node centrality in the graph to identify critical threats.
        """
        if self.graph is None:
            self.output.log("Graph not available for centrality analysis")
            return {}
        
        self.output.log("\n=== CENTRALITY ANALYSIS ===")
        
        centrality_measures = {}
        
        try:
            # Degree Centrality (always available)
            degree_centrality = nx.degree_centrality(self.graph)
            in_degree_centrality = nx.in_degree_centrality(self.graph)
            out_degree_centrality = nx.out_degree_centrality(self.graph)
            
            centrality_measures['degree'] = degree_centrality
            centrality_measures['in_degree'] = in_degree_centrality
            centrality_measures['out_degree'] = out_degree_centrality
            
            # Betweenness Centrality (always available but can be slow)
            self.output.log("Calculating betweenness centrality...")
            betweenness_centrality = nx.betweenness_centrality(self.graph)
            centrality_measures['betweenness'] = betweenness_centrality
            
            # Closeness Centrality (always available)
            self.output.log("Calculating closeness centrality...")
            closeness_centrality = nx.closeness_centrality(self.graph)
            centrality_measures['closeness'] = closeness_centrality
            
            # PageRank (always available)
            self.output.log("Calculating PageRank...")
            pagerank = nx.pagerank(self.graph)
            centrality_measures['pagerank'] = pagerank
            
            # Eigenvector Centrality (requires scipy for better convergence)
            if HAS_SCIPY:
                try:
                    self.output.log("Calculating eigenvector centrality...")
                    eigenvector_centrality = nx.eigenvector_centrality(self.graph, max_iter=1000)
                    centrality_measures['eigenvector'] = eigenvector_centrality
                except:
                    self.output.log("⚠️  Eigenvector centrality not calculable (graph might not be strongly connected)")
            
        except Exception as e:
            self.output.log(f"Error calculating centrality measures: {e}")
            return {}
        
        # Show results
        self._display_centrality_results(centrality_measures)
        
        return centrality_measures
    def _display_centrality_results(self, centrality_measures):
        """Displays centrality measure results."""
        if not centrality_measures:
            return
        
        # Use the configurable parameter for the number of nodes
        top_n = ANALYSIS_PARAMETERS["top_centrality_nodes"]
        
        self.output.log(f"\n🎯 MOST CENTRAL NODES (TOP {top_n} for each measure):")
        
        for measure_name, measure_values in centrality_measures.items():
            self.output.log(f"\n--- {measure_name.upper()} CENTRALITY ---")
            
            # Sort by centrality value
            sorted_nodes = sorted(measure_values.items(), key=lambda x: x[1], reverse=True)[:top_n]
            
            for i, (node, centrality) in enumerate(sorted_nodes, 1):
                # Get node category
                if self.graph and node in self.graph.nodes:
                    category = self.graph.nodes[node].get('category', 'Unknown')
                    self.output.log(f"  {i}. [{category}] {node}: {centrality:.4f}")
                else:
                    self.output.log(f"  {i}. {node}: {centrality:.4f}")
                    self.output.log(f"  {i}. {node}: {centrality:.4f}")
    
    def analyze_critical_paths(self, max_paths_per_pair=None, max_length=None):
        """
        Identifies and analyzes the most critical attack paths.
        
        Args:
            max_paths_per_pair (int): Maximum number of paths to analyze per source-target pair
            max_length (int): Maximum length of paths to consider
        """
        if self.graph is None:
            self.output.log("Graph not available for critical path analysis")
            return []
        
        # Use configurable parameters if not specified        
        if max_paths_per_pair is None:
            max_paths_per_pair = ANALYSIS_PARAMETERS["max_paths_per_pair"]
        if max_length is None:
            max_length = ANALYSIS_PARAMETERS["max_critical_path_length"]
        
        self.output.log(f"\n=== CRITICAL PATH ANALYSIS ===")
        self.output.log(f"Parameters: max_paths_per_pair={max_paths_per_pair}, max_length={max_length}")
        
        # Get high-risk threats for analysis once
        high_risk_threats = self._get_top_risk_threats(top_n=10)
        
        # Identify critical source and target threats
        critical_sources = self._identify_critical_sources()
        critical_targets = self._identify_critical_targets()
        
        # Remove duplicates and limit the number for performance
        critical_sources = list(set(critical_sources))[:10]  # Max 10 sources
        critical_targets = list(set(critical_targets))[:10]   # Max 10 targets
        
        self.output.log(f"\nCritical source threats identified: {len(critical_sources)}")
        self.output.log(f"Critical target threats identified: {len(critical_targets)}")
        
        # For the subset, we analyze all the most interesting combinations
        critical_paths = []
        analyzed_pairs = 0
        max_pairs = min(len(critical_sources) * len(critical_targets), 25)  # Reduced for performance
        
        # Use a set to avoid analyzing the same pair multiple times
        analyzed_combinations = set()
        
        for source in critical_sources:
            for target in critical_targets:
                combination = (source, target)
                if (source != target and 
                    analyzed_pairs < max_pairs and 
                    combination not in analyzed_combinations):
                    
                    analyzed_combinations.add(combination)
                    analyzed_pairs += 1
                    try:
                        # Find all simple paths
                        paths = list(nx.all_simple_paths(self.graph, source, target, cutoff=max_length))
                        
                        # Take only the first N paths for performance and avoid duplicates
                        unique_paths = []
                        for path in paths[:max_paths_per_pair]:                            
                            path_tuple = tuple(path)
                            if path_tuple not in [tuple(p['path']) for p in unique_paths]:
                                score = self._calculate_path_criticality(path, high_risk_threats)
                                unique_paths.append({
                                    'path': path,
                                    'source': source,
                                    'target': target,
                                    'length': len(path),
                                    'score': score
                                })
                        
                        critical_paths.extend(unique_paths)
                    except nx.NetworkXNoPath:
                        continue
                    except Exception as e:
                        self.output.log(f"Error calculating paths {source} -> {target}: {e}")
                        continue        
        # Remove duplicate paths based on the actual path
        unique_critical_paths = []
        seen_paths = set()
        
        for path_info in critical_paths:
            path_tuple = tuple(path_info['path'])
            if path_tuple not in seen_paths:
                seen_paths.add(path_tuple)
                unique_critical_paths.append(path_info)
        
        # Sort by criticality
        unique_critical_paths.sort(key=lambda x: x['score'], reverse=True)        
        self.output.log(f"\nTotal critical paths analyzed: {len(critical_paths)}")
        self.output.log(f"Unique paths after deduplication: {len(unique_critical_paths)}")
        self.output.log(f"Source-target pairs analyzed: {analyzed_pairs}")
        
        # Show results
        top_paths = ANALYSIS_PARAMETERS["top_critical_paths"]
        self._display_critical_paths(unique_critical_paths[:top_paths])
        
        return unique_critical_paths
    
    def _get_top_impact_threats(self, top_n=10):
        """Gets the top N threats with the highest impact from the configured THREAT_FILE_NAME file."""
        # Use the subset file path that was configured at initialization
        subset_file = self.subset_file_path
        
        if not os.path.exists(subset_file):
            self.output.log(f"⚠️  File {subset_file} not found, using predefined keywords.")
            return []
        
        try:
            # Read the configured threat file
            df = pd.read_csv(subset_file, sep=';')
            
            # Check that the Impact column exists
            if 'Impact' not in df.columns:
                self.output.log(f"⚠️  'Impact' column not found in {THREAT_FILE_NAME}. Available columns: {list(df.columns)}")
                return []
            
            # Define mapping of impact values to numbers for sorting
            impact_mapping = {
                'Very Low': 1,
                'Low': 2, 
                'Medium': 3,
                'High': 4,
                'Very High': 5
            }
            
            # Convert impact values to numbers
            df['Impact_Score'] = df['Impact'].map(impact_mapping)
            
            # Remove rows with unrecognized Impact values
            df = df.dropna(subset=['Impact_Score'])
            
            if len(df) == 0:
                self.output.log(f"⚠️  No threats with valid Impact values found in {THREAT_FILE_NAME}")
                return []
            
            # Sort by Impact_Score in descending order and take the top N
            top_threats = df.nlargest(top_n, 'Impact_Score')
            
            # Return only threat names
            threat_names = top_threats['THREAT'].tolist()
            
            self.output.log(f"📊 Top {len(threat_names)} threats with highest impact:")
            for i, threat in enumerate(threat_names, 1):
                impact_value = top_threats.iloc[i-1]['Impact']
                self.output.log(f"   {i:2d}. {threat} (Impact: {impact_value})")
            
            return threat_names
        except Exception as e:
            self.output.log(f"⚠️  Error reading {THREAT_FILE_NAME}: {e}")
            return []

    def _get_top_likelihood_threats(self, top_n=10):
        """Gets threats with highest Likelihood from the configured THREAT_FILE_NAME file"""
        try:
            # Read the configured threat file
            df = pd.read_csv(self.subset_file_path, sep=';')
            
            # Check that the Likelihood column exists
            if 'Likelihood' not in df.columns:
                self.output.log(f"⚠️  'Likelihood' column not found in {THREAT_FILE_NAME}. Using fallback.")
                return [
                    'Social Engineering', 'Unauthorized access', 'Physical access',
                    'Supply Chain', 'Legacy Software', 'Malicious code'
                ]
            
            # Define mapping of likelihood values to numbers for sorting
            likelihood_mapping = {
                'Very Low': 1,
                'Low': 2, 
                'Medium': 3,
                'High': 4,
                'Very High': 5
            }
            
            # Convert likelihood values to numbers
            df['Likelihood_Score'] = df['Likelihood'].map(likelihood_mapping)
            
            # Remove rows with unrecognized Likelihood values
            df = df.dropna(subset=['Likelihood_Score'])
            
            if len(df) == 0:
                self.output.log(f"⚠️  No threats with valid Likelihood values found. Using fallback.")
                return [
                    'Social Engineering', 'Unauthorized access', 'Physical access',
                    'Supply Chain', 'Legacy Software', 'Malicious code'
                ]
            
            # Sort by Likelihood_Score in descending order and take the top N
            top_threats = df.nlargest(top_n, 'Likelihood_Score')
            
            # Return only threat names
            threat_names = top_threats['THREAT'].tolist()
            
            self.output.log(f"📊 Top {len(threat_names)} threats with highest likelihood:")
            for i, threat in enumerate(threat_names, 1):
                likelihood_value = top_threats.iloc[i-1]['Likelihood']
                self.output.log(f"   {i:2d}. {threat} (Likelihood: {likelihood_value})")
            
            return threat_names
            
        except Exception as e:
            self.output.log(f"⚠️  Error reading threats with high Likelihood: {e}")
            # Fallback to hardcoded list
            return [
                'Social Engineering', 'Unauthorized access', 'Physical access',
                'Supply Chain', 'Legacy Software', 'Malicious code'
            ]

    def _get_top_risk_threats(self, top_n=10):
        """Gets threats with highest Risk from the configured THREAT_FILE_NAME file"""
        try:
            # Read the configured threat file
            df = pd.read_csv(self.subset_file_path, sep=';')
            
            # Check that the Risk column exists
            if 'Risk' not in df.columns:
                self.output.log(f"⚠️  'Risk' column not found in {THREAT_FILE_NAME}. Using fallback.")
                return [
                    'Seizure', 'Control', 'Satellite', 'Destruction', 'Failure',
                    'Security', 'Unauthorized', 'Malicious', 'Denial'
                ]
            
            # Define mapping of risk values to numbers for sorting
            risk_mapping = {
                'Very Low': 1,
                'Low': 2, 
                'Medium': 3,
                'High': 4,
                'Very High': 5
            }
            
            # Convert risk values to numbers
            df['Risk_Score'] = df['Risk'].map(risk_mapping)
            
            # Remove rows with unrecognized Risk values
            df = df.dropna(subset=['Risk_Score'])
            
            if len(df) == 0:
                self.output.log(f"⚠️  No threats with valid Risk values found. Using fallback.")
                return [
                    'Seizure', 'Control', 'Satellite', 'Destruction', 'Failure',
                    'Security', 'Unauthorized', 'Malicious', 'Denial'
                ]
            
            # Sort by Risk_Score in descending order and take the top N
            top_threats = df.nlargest(top_n, 'Risk_Score')
            
            # Return only threat names
            threat_names = top_threats['THREAT'].tolist()
            
            self.output.log(f"📊 Top {len(threat_names)} threats with highest risk:")
            for i, threat in enumerate(threat_names, 1):
                risk_value = top_threats.iloc[i-1]['Risk']
                self.output.log(f"   {i:2d}. {threat} (Risk: {risk_value})")
            
            return threat_names
            
        except Exception as e:
            self.output.log(f"⚠️  Error reading threats with high Risk: {e}")
            # Fallback to hardcoded list
            return [
                'Seizure', 'Control', 'Satellite', 'Destruction', 'Failure',
                'Security', 'Unauthorized', 'Malicious', 'Denial'
            ]

    def _identify_critical_targets(self):
        """Identifies critical threat targets based on in-degree and category."""
        if self.graph is None:
            return []
            
        in_degrees = dict(self.graph.in_degree())
        
        # Define critical categories for space systems
        critical_categories = {'NAA', 'EIH', 'PA'}  # Nefarious, Eavesdropping, Physical Access
        
        # Get threats with highest impact from the configured THREAT_FILE_NAME file
        critical_keywords = self._get_top_impact_threats(top_n=10)
        
        # Fallback keywords if unable to read from file
        if not critical_keywords:
            critical_keywords = [
                'Seizure of control', 'Denial of Service', 'Data modification',
                'Firmware corruption', 'Satellite bus', 'Compromising',
                'Destruction', 'Failure of power', 'Security services failure'
            ]
        
        critical_targets = []
        
        for node in self.graph.nodes():
            score = in_degrees.get(node, 0)
            
            # Bonus for critical category
            node_category = self.graph.nodes[node].get('category', '')
            if node_category in critical_categories:
                score += 2
            
            # Bonus for critical keywords
            for keyword in critical_keywords:
                if keyword.lower() in node.lower():
                    score += 3
                    break
            
            if score >= 2:  # Minimum threshold
                critical_targets.append((node, score))
          # Sort by score and return only nodes
        critical_targets.sort(key=lambda x: x[1], reverse=True)
        return [node for node, score in critical_targets]
    
    def _identify_critical_sources(self):
        """Identifies critical threat sources based on out-degree and type."""
        if self.graph is None:
            return []
            
        out_degrees = dict(self.graph.out_degree())
        
        # Get threats with highest likelihood from the configured THREAT_FILE_NAME file
        initial_threat_keywords = self._get_top_likelihood_threats(top_n=10)
        
        # Fallback keywords if unable to read from file
        if not initial_threat_keywords:
            initial_threat_keywords = [
                'Social Engineering', 'Unauthorized access', 'Physical access',
                'Supply Chain', 'Legacy Software', 'Malicious code'
            ]

        critical_sources = []
        
        for node in self.graph.nodes():
            score = out_degrees.get(node, 0)
            
            # Bonus for typical initial threats
            for keyword in initial_threat_keywords:
                if keyword.lower() in node.lower():
                    score += 2
                    break
            
            if score >= 1:  # Lower threshold for sources
                critical_sources.append((node, score))
        
        # Sort by score and return only nodes        
        critical_sources.sort(key=lambda x: x[1], reverse=True)
        return [node for node, score in critical_sources]
    
    def _calculate_path_criticality(self, path, high_risk_threats=None):
        """
        Calculate a criticality score for an attack path.
        
        Args:
            path (list): List of nodes that form the path
            high_risk_threats (list): List of high-risk threats (to avoid multiple calls)
            
        Returns:
            float: Criticality score
        """
        if len(path) < 2 or self.graph is None:
            return 0
        
        score = 0
        
        # Criticality factors:
        # 1. Path length (longer paths are more complex but also more informative)
        length_factor = len(path) * 0.5
        
        # 2. Types of relations in the path
        relation_weights = {
            'Enables': 3,
            'Causes': 4,
            'Leads-to': 2,
            'Triggers': 3,
            'Amplifies': 2
        }
        
        relation_score = 0
        for i in range(len(path) - 1):
            edge_data = self.graph[path[i]][path[i+1]]
            relation_type = edge_data.get('relation_type', 'Unknown')
            relation_score += relation_weights.get(relation_type, 1)        
        # 3. Criticality of nodes in the path
        node_criticality = 0
        
        # Use the high-risk threats passed as parameter or get them if not provided
        if high_risk_threats is None:
            critical_threats = self._get_top_risk_threats(top_n=10)
        else:
            critical_threats = high_risk_threats
        
        for node in path:
            # Check if the node corresponds to one of the high-risk threats
            for threat in critical_threats:
                if threat.lower() in node.lower():
                    node_criticality += 1
                    break
        
        # 4. Diversity of categories traversed
        categories = set()
        for node in path:
            category = self.graph.nodes[node].get('category', 'Unknown')
            categories.add(category)
        
        category_diversity = len(categories) * 0.5
        
        # Final calculation
        score = length_factor + relation_score + node_criticality + category_diversity
        
        return score
    
    def _display_critical_paths(self, critical_paths):
        """Display critical paths in a formatted way."""
        
        if not critical_paths or self.graph is None:
            self.output.log("No critical paths identified.")
            return
        
        self.output.log(f"\n🚨 TOP {len(critical_paths)} CRITICAL PATHS IDENTIFIED:")
        
        for i, path_info in enumerate(critical_paths, 1):
            path = path_info['path']
            score = path_info['score']
            length = path_info['length']

            danger = (score - 2) / (48)
            danger = min(max(danger, 0), 1) 

            self.output.log(f"\n🔥 CRITICAL PATH #{i} (Score: {score:.2f}, Danger: {danger:.2f}, Length: {length})")
            self.output.log(f"   From: {path[0]}")
            self.output.log(f"   To:   {path[-1]}")
            self.output.log("   Sequence:")
            
            for j in range(len(path) - 1):
                edge_data = self.graph[path[j]][path[j+1]]
                relation = edge_data.get('relation_type', 'Unknown')
                source_cat = self.graph.nodes[path[j]].get('category', '?')
                target_cat = self.graph.nodes[path[j+1]].get('category', '?')
                
                self.output.log(f"     {j+1}. [{source_cat}] {path[j]}")
                self.output.log(f"        --({relation})--> [{target_cat}] {path[j+1]}")
    
    def analyze_attack_surface(self):
        """
        Analyze the attack surface by identifying entry points and final targets.
        """
        if self.graph is None:
            self.output.log("Graph not available")
            return {}
        
        self.output.log("\n=== ATTACK SURFACE ANALYSIS ===")
        
        # Entry points (nodes with low in-degree but high out-degree)
        in_degrees = dict(self.graph.in_degree())
        out_degrees = dict(self.graph.out_degree())
        
        entry_points = []
        final_targets = []
        
        for node in self.graph.nodes():
            in_deg = in_degrees[node]
            out_deg = out_degrees[node]
            
            # Entry points: few inputs, many outputs
            if in_deg <= 1 and out_deg >= 3:
                entry_points.append((node, out_deg))
            
            # Final targets: many inputs, few outputs
            if in_deg >= 3 and out_deg <= 1:
                final_targets.append((node, in_deg))
        
        # Sort by relevance
        entry_points.sort(key=lambda x: x[1], reverse=True)
        final_targets.sort(key=lambda x: x[1], reverse=True)
        
        self.output.log(f"\n🚪 ENTRY POINTS IDENTIFIED ({len(entry_points)}):")
        for node, out_deg in entry_points[:10]:
            category = self.graph.nodes[node].get('category', '?')
            self.output.log(f"  [{category}] {node} (enables {out_deg} attacks)")
        
        self.output.log(f"\n🎯 FINAL TARGETS IDENTIFIED ({len(final_targets)}):")
        for node, in_deg in final_targets[:10]:
            category = self.graph.nodes[node].get('category', '?')
            self.output.log(f"  [{category}] {node} (receives {in_deg} attacks)")
        
        return {
            'entry_points': entry_points,
            'final_targets': final_targets
        }
    
    def analyze_threat_connections(self, target_threat=None):
        """
        Analyze the connections of a specific threat in the graph.
        Shows predecessors, successors and paths that involve the threat.
        
        Args:
            target_threat (str): Name of the threat to analyze. If None, uses the one configured in STAR_GRAPH_CONFIG
        """
        if self.graph is None:
            self.output.log("Graph not available for connection analysis")
            return {}
        
        # Use the configured threat if not specified
        if target_threat is None:
            target_threat = STAR_GRAPH_CONFIG.get("center_threat", "Social Engineering")
        
        if target_threat not in self.graph.nodes():
            self.output.log(f"⚠️ Threat '{target_threat}' not found in graph")
            
            # Suggest similar threats
            available_threats = list(self.graph.nodes())
            similar_threats = [t for t in available_threats if target_threat.lower() in t.lower() or t.lower() in target_threat.lower()]
            
            if similar_threats:
                self.output.log(f"💡 Similar threats available: {similar_threats[:5]}")
            else:
                self.output.log(f"💡 Some available threats: {available_threats[:10]}")
            return {}
        
        self.output.log(f"\n🔍 CONNECTION ANALYSIS FOR: '{target_threat}'")
        self.output.log("=" * 70)
        
        # Base node information
        category = self.graph.nodes[target_threat].get('category', 'Unknown')
        in_degree = self.graph.in_degree(target_threat)
        out_degree = self.graph.out_degree(target_threat)
        total_degree = in_degree + out_degree
        
        self.output.log(f"📊 BASIC INFORMATION:")
        self.output.log(f"   Category: {category}")
        self.output.log(f"   Incoming connections: {in_degree}")
        self.output.log(f"   Outgoing connections: {out_degree}")
        self.output.log(f"   Total connections: {total_degree}")
        
        # Analysis of predecessors (threats that lead to this one)
        predecessors = list(self.graph.predecessors(target_threat))
        self.output.log(f"\n🔽 PREDECESSORS ({len(predecessors)}) - Threats that LEAD TO '{target_threat}':")
        
        if predecessors:
            # Sort by relevance (nodes with more outgoing connections are more critical)
            pred_scores = [(pred, self.graph.out_degree(pred)) for pred in predecessors]
            pred_scores.sort(key=lambda x: x[1], reverse=True)
            
            for i, (pred, out_deg) in enumerate(pred_scores, 1):
                pred_category = self.graph.nodes[pred].get('category', '?')
                edge_data = self.graph[pred][target_threat]
                relation_type = edge_data.get('relation_type', 'Unknown')
                
                self.output.log(f"   {i:2d}. [{pred_category}] {pred}")
                self.output.log(f"       --({relation_type})--> {target_threat}")
                self.output.log(f"       (out-degree: {out_deg})")
        else:
            self.output.log(f"   ⚠️ No predecessors found. '{target_threat}' might be an entry point.")
        
        # Analysis of successors (threats enabled by this one)
        successors = list(self.graph.successors(target_threat))
        self.output.log(f"\n🔼 SUCCESSORS ({len(successors)}) - Threats ENABLED BY '{target_threat}':")
        
        if successors:
            # Sort by relevance (nodes with more incoming connections are more critical targets)
            succ_scores = [(succ, self.graph.in_degree(succ)) for succ in successors]
            succ_scores.sort(key=lambda x: x[1], reverse=True)
            
            for i, (succ, in_deg) in enumerate(succ_scores, 1):
                succ_category = self.graph.nodes[succ].get('category', '?')
                edge_data = self.graph[target_threat][succ]
                relation_type = edge_data.get('relation_type', 'Unknown')
                
                self.output.log(f"   {i:2d}. [{succ_category}] {succ}")
                self.output.log(f"       {target_threat} --({relation_type})-->")
                self.output.log(f"       (in-degree: {in_deg})")
        else:
            self.output.log(f"   ⚠️ No successors found. '{target_threat}' might be an end point.")
        
        # Analysis of paths that traverse this threat
        self._analyze_paths_through_threat(target_threat)
        
        # Specific centrality analysis for this node
        self._analyze_threat_centrality(target_threat)
        
        # Analysis of second-level neighbors
        self._analyze_second_level_neighbors(target_threat)
        
        # Save connection visualization if requested
        if THREAT_CONNECTION_ANALYSIS.get("save_visualization", False):
            self._save_threat_connection_visualization(target_threat, predecessors, successors)
        
        return {
            'threat': target_threat,
            'category': category,
            'in_degree': in_degree,
            'out_degree': out_degree,
            'predecessors': predecessors,
            'successors': successors
        }
    
    def _analyze_paths_through_threat(self, target_threat, max_paths=5):
        """Analyzes attack paths that pass through the specified threat"""
        if self.graph is None:
            self.output.log(f"\n🛤️ PATHS THROUGH '{target_threat}': Graph not available")
            return
            
        self.output.log(f"\n🛤️ PATHS THROUGH '{target_threat}':")
        
        # Find all possible entry points (nodes with low in-degree)
        entry_points = [node for node in self.graph.nodes() 
                       if self.graph.in_degree(node) <= 1 and node != target_threat]
        
        # Find all possible final targets (nodes with low out-degree)
        final_targets = [node for node in self.graph.nodes() 
                        if self.graph.out_degree(node) <= 1 and node != target_threat]
        
        paths_found = 0
        max_total_paths = max_paths * 2  # Limit total number for performance
        
        self.output.log(f"   Searching paths from {len(entry_points)} entry points to {len(final_targets)} final targets...")
        
        for entry in entry_points[:10]:  # Limit entry points for performance
            if paths_found >= max_total_paths:
                break
                
            for target in final_targets[:10]:  # Limit targets for performance
                if paths_found >= max_total_paths:
                    break
                    
                try:
                    # Search for paths that pass through target_threat
                    paths_to_threat = list(nx.all_simple_paths(self.graph, entry, target_threat, cutoff=4))
                    paths_from_threat = list(nx.all_simple_paths(self.graph, target_threat, target, cutoff=4))
                    
                    # Combine paths
                    for path_to in paths_to_threat[:2]:  # Max 2 paths per combination
                        for path_from in paths_from_threat[:2]:
                            if paths_found >= max_paths:
                                break
                            
                            # Combine paths removing the duplicate target_threat node
                            full_path = path_to + path_from[1:]
                            
                            if len(full_path) >= 3:  # Significant paths
                                paths_found += 1
                                self.output.log(f"\n   📍 PATH #{paths_found}:")
                                self.output.log(f"     {' → '.join(full_path)}")
                                self.output.log(f"     Length: {len(full_path)} nodes")
                                
                                # Show relations
                                for i in range(len(full_path) - 1):
                                    if self.graph.has_edge(full_path[i], full_path[i + 1]):
                                        edge_data = self.graph[full_path[i]][full_path[i + 1]]
                                        relation = edge_data.get('relation_type', 'Unknown')
                                        self.output.log(f"       {full_path[i]} --({relation})-> {full_path[i + 1]}")
                        
                        if paths_found >= max_paths:
                            break
                    
                except (nx.NetworkXNoPath, nx.NetworkXError):
                    continue
                except Exception as e:
                    self.output.log(f"     ⚠️ Error calculating paths: {e}")
                    continue
        
        if paths_found == 0:
            self.output.log(f"   ⚠️ No significant paths found that traverse '{target_threat}'")
        else:
            self.output.log(f"\n   ✅ Found {paths_found} paths that traverse '{target_threat}'")
    
    def _analyze_threat_centrality(self, target_threat):
        """Analyzes specific centrality measures for the threat"""
        if self.graph is None:
            self.output.log(f"\n📈 CENTRALITY MEASURES FOR '{target_threat}': Graph not available")
            return
            
        self.output.log(f"\n📈 CENTRALITY MEASURES FOR '{target_threat}':")
        
        try:
            # Degree centrality
            degree_cent = nx.degree_centrality(self.graph)[target_threat]
            in_degree_cent = nx.in_degree_centrality(self.graph)[target_threat]
            out_degree_cent = nx.out_degree_centrality(self.graph)[target_threat]
            
            self.output.log(f"   Degree centrality: {degree_cent:.4f}")
            self.output.log(f"   In-degree centrality: {in_degree_cent:.4f}")
            self.output.log(f"   Out-degree centrality: {out_degree_cent:.4f}")
            
            # Betweenness centrality
            betweenness_cent = nx.betweenness_centrality(self.graph)[target_threat]
            self.output.log(f"   Betweenness centrality: {betweenness_cent:.4f}")
            
            # Closeness centrality
            closeness_cent = nx.closeness_centrality(self.graph)[target_threat]
            self.output.log(f"   Closeness centrality: {closeness_cent:.4f}")
            
            # PageRank
            pagerank = nx.pagerank(self.graph)[target_threat]
            self.output.log(f"   PageRank: {pagerank:.4f}")
            
            # Interpretation
            total_nodes = self.graph.number_of_nodes()
            self.output.log(f"\n   💡 INTERPRETATION:")
            
            if degree_cent > 0.1:
                self.output.log(f"     - High connectivity: the threat is well connected in the network")
            elif degree_cent < 0.05:
                self.output.log(f"     - Low connectivity: the threat has few direct connections")
            
            if betweenness_cent > 0.1:
                self.output.log(f"     - High control: the threat is an important bridge between other threats")
            elif betweenness_cent < 0.01:
                self.output.log(f"     - Low control: the threat rarely acts as an intermediary")
            
            if pagerank > 1.0 / total_nodes * 2:
                self.output.log(f"     - High importance: the threat is considered important by the network")
            
        except Exception as e:
            self.output.log(f"   ⚠️ Error calculating centrality: {e}")
    
    def _analyze_second_level_neighbors(self, target_threat):
        """Analyzes second-level neighbors (neighbors of neighbors)"""
        if self.graph is None:
            self.output.log(f"\n🔍 SECOND-LEVEL NEIGHBORS FOR '{target_threat}': Graph not available")
            return
            
        self.output.log(f"\n🔍 SECOND-LEVEL NEIGHBORS FOR '{target_threat}':")
        
        # Direct neighbors
        direct_neighbors = set(self.graph.predecessors(target_threat)) | set(self.graph.successors(target_threat))
        
        # Second-level neighbors
        second_level = set()
        for neighbor in direct_neighbors:
            second_level.update(self.graph.predecessors(neighbor))
            second_level.update(self.graph.successors(neighbor))
        
        # Remove the node itself and direct neighbors
        second_level.discard(target_threat)
        second_level -= direct_neighbors
        
        self.output.log(f"   Direct neighbors: {len(direct_neighbors)}")
        self.output.log(f"   Second-level neighbors: {len(second_level)}")
        
        if second_level:
            # Sort by relevance (sum of in_degree and out_degree)
            in_degrees = dict(self.graph.in_degree())
            out_degrees = dict(self.graph.out_degree())
            second_level_scores = [(node, in_degrees.get(node, 0) + out_degrees.get(node, 0)) for node in second_level]
            second_level_scores.sort(key=lambda x: x[1], reverse=True)
            
            self.output.log(f"\n   🎯 TOP SECOND-LEVEL NEIGHBORS (by connectivity):")
            for i, (node, degree) in enumerate(second_level_scores[:10], 1):
                category = self.graph.nodes[node].get('category', '?')
                self.output.log(f"     {i:2d}. [{category}] {node} (degree: {degree})")
        else:
            self.output.log(f"   ⚠️ No second-level neighbors found")

    def _save_threat_connection_visualization(self, target_threat, predecessors, successors):
        """
        Saves a visualization of the connection graph for a specific threat.
        
        Args:
            target_threat (str): The central threat
            predecessors (list): List of predecessors
            successors (list): List of successors
        """
        if self.graph is None:
            self.output.log(f"\n💾 UNABLE TO SAVE VISUALIZATION: graph not available for '{target_threat}'")
            return
            
        try:
            self.output.log(f"\n💾 SAVING CONNECTION VISUALIZATION FOR '{target_threat}'...")
            
            # Create a subgraph with the central threat and its connections
            nodes_to_include = {target_threat}
            nodes_to_include.update(predecessors)
            nodes_to_include.update(successors)
            
            # Add second-level neighbors if configured
            max_distance = THREAT_CONNECTION_ANALYSIS.get("max_distance", 2)
            if max_distance >= 2:
                for neighbor in list(predecessors) + list(successors):
                    if THREAT_CONNECTION_ANALYSIS.get("include_predecessors", True):
                        nodes_to_include.update(self.graph.predecessors(neighbor))
                    if THREAT_CONNECTION_ANALYSIS.get("include_successors", True):
                        nodes_to_include.update(self.graph.successors(neighbor))
            
            # Create the subgraph
            subgraph = self.graph.subgraph(nodes_to_include).copy()
            
            # Find nodes that are both predecessors and successors
            both_pred_and_succ = set(predecessors) & set(successors)
            
            # Add duplicate nodes and edges for nodes that are both predecessors and successors
            for node in both_pred_and_succ:
                if node in successors and node in predecessors:
                    # Create duplicate node name
                    duplicate_node_name = f"{node}_successor_copy"
                    
                    # Add the duplicate node to the subgraph
                    subgraph.add_node(duplicate_node_name, **self.graph.nodes[node])
                    
                    # Add edge from central threat to duplicate node (successor relationship)
                    if self.graph.has_edge(target_threat, node):
                        edge_data = self.graph[target_threat][node]
                        subgraph.add_edge(target_threat, duplicate_node_name, **edge_data)
                    
                    # Add edges from duplicate node to its successors
                    for successor in self.graph.successors(node):
                        if successor in nodes_to_include:
                            if self.graph.has_edge(node, successor):
                                edge_data = self.graph[node][successor]
                                subgraph.add_edge(duplicate_node_name, successor, **edge_data)
            
            if len(subgraph.nodes()) == 0:
                self.output.log("   ⚠️ No nodes to visualize")
                return
            
            # Configure the visualization
            plt.figure(figsize=(16, 12))
            plt.clf()

            # Hierarchical layout: central threat in center, predecessors left, successors right
            pos = self._create_hierarchical_threat_connections_layout(
                subgraph, target_threat, predecessors, successors
            )

            # Colors and sizes for different types of nodes - handle duplicates properly
            node_colors = []
            node_sizes = []
            
            # Process all nodes in the subgraph (including duplicates)
            for node in subgraph.nodes():
                if node == target_threat:
                    node_colors.append('#FF4444')  # Red for the central threat
                    node_sizes.append(2500)
                elif node.endswith('_successor_copy'):
                    # This is a duplicate node representing a successor
                    node_colors.append('#44FF44')  # Green for successors
                    node_sizes.append(1500)
                elif node in predecessors:
                    node_colors.append('#4444FF')  # Blue for predecessors
                    node_sizes.append(1500)
                elif node in successors:
                    node_colors.append('#44FF44')  # Green for successors
                    node_sizes.append(1500)
                else:
                    node_colors.append('#FFAA44')  # Orange for second-level neighbors
                    node_sizes.append(1000)

            # Draw all nodes with their assigned colors using networkx
            nx.draw_networkx_nodes(subgraph, pos,
                                 node_color=node_colors,  # type: ignore
                                 node_size=node_sizes,  # type: ignore
                                 alpha=0.8,
                                 edgecolors='black',
                                 linewidths=2)

            # Draw edges - only direct connections to/from the central threat
            direct_edges = []
            edge_colors = []
            
            for edge in subgraph.edges():
                source, target = edge
                # Only include edges that involve the central threat
                if source == target_threat or target == target_threat:
                    direct_edges.append(edge)
                    edge_colors.append('#333333')  # Dark gray for direct connections
            
            if direct_edges:
                nx.draw_networkx_edges(subgraph, pos,
                                     edgelist=direct_edges,
                                     edge_color=edge_colors,  # type: ignore
                                     width=2,
                                     alpha=0.7,
                                     arrows=True,
                                     arrowsize=20,
                                     arrowstyle='->')
            
            # Draw node labels - simplified since all nodes are now in subgraph
            labels = {}
            for node in subgraph.nodes():
                if node.endswith('_successor_copy'):
                    # Show original name for duplicate nodes
                    original_name = node.replace('_successor_copy', '')
                    labels[node] = original_name
                else:
                    labels[node] = node
            
            nx.draw_networkx_labels(subgraph, pos, labels,
                                  font_size=9,
                                  font_weight='bold',
                                  font_color='white',
                                  bbox=dict(boxstyle='round,pad=0.3', 
                                          facecolor='black', 
                                          alpha=0.7))

            # Title and legend
            plt.title(f"Threat Connections: {target_threat}", 
                     fontsize=16, fontweight='bold', pad=20)

            # Create simplified legend
            legend_elements = [
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FF4444', 
                          markersize=15, label='Central Threat'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#4444FF', 
                          markersize=12, label='Predecessors (left)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#44FF44', 
                          markersize=12, label='Successors (right)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FFAA44', 
                          markersize=10, label='Second Level'),
                Line2D([0], [0], color='#333333', linewidth=2, label='Direct Connection')
            ]
            
            plt.legend(handles=legend_elements, loc='upper right', bbox_to_anchor=(1.15, 1))
            
            # Simplified additional info (no duplicates)
            info_text = (f"Total Nodes: {len(subgraph.nodes())}\n"
                        f"Total Edges: {len(subgraph.edges())}\n"
                        f"Predecessors: {len(predecessors)}\n"
                        f"Successors: {len(successors)}")

            plt.text(0.02, 0.98, info_text, transform=plt.gca().transAxes,
                    fontsize=10, verticalalignment='top',
                    bbox=dict(boxstyle='round,pad=0.5', facecolor='lightgray', alpha=0.8))
            
            plt.axis('off')
            plt.tight_layout()
            
            # Save the image
            # Remove invalid characters from the filename
            safe_threat_name = "".join(c for c in target_threat if c.isalnum() or c in (' ', '-', '_')).rstrip()
            safe_threat_name = safe_threat_name.replace(' ', '_')
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"threat_connections_{safe_threat_name}_{timestamp}.png"
            
            # Create Output directory if it doesn't exist
            output_dir = os.path.join(get_output_path(), "Output")
            os.makedirs(output_dir, exist_ok=True)
            filepath = os.path.join(output_dir, filename)
            
            plt.savefig(filepath, dpi=300, bbox_inches='tight', 
                       facecolor='white', edgecolor='none')

            self.output.log(f"   ✅ Visualization saved as: {filepath}")
            self.output.log(f"   📊 Total Nodes: {len(subgraph.nodes())}, Total Edges: {len(subgraph.edges())}")

            plt.close()
            
        except Exception as e:
            self.output.log(f"   ❌ Error saving visualization: {e}")
            import traceback
            self.output.log(f"   Error details: {traceback.format_exc()}")

    def find_attack_paths(self, source_threat, target_threat, max_length=5):
        """
        Find all attack paths between two threats.
        
        Args:
            source_threat (str): Starting threat
            target_threat (str): Destination threat
            max_length (int): Maximum path length
        """
        if self.graph is None:
            self.output.log("Graph not available")
            return []
        
        if source_threat not in self.graph.nodes():
            self.output.log(f"Threat '{source_threat}' not found in graph")
            return []
        
        if target_threat not in self.graph.nodes():
            self.output.log(f"Threat '{target_threat}' not found in graph")
            return []
        
        try:
            paths = list(nx.all_simple_paths(self.graph, source_threat, target_threat, cutoff=max_length))
            
            # Check for direct connection (path of length 1)
            if self.graph.has_edge(source_threat, target_threat):
                # Add direct path if not already included
                direct_path = [source_threat, target_threat]
                if direct_path not in paths:
                    paths.insert(0, direct_path)  # Put direct path first
            
            self.output.log(f"\n=== ATTACK PATHS: {source_threat} → {target_threat} ===")
            if not paths:
                self.output.log("No paths found")
            else:
                for i, path in enumerate(paths, 1):
                    self.output.log(f"\nPath {i} (length {len(path)-1}):")
                    for j in range(len(path)-1):
                        edge_data = self.graph[path[j]][path[j+1]]
                        relation = edge_data.get('relation_type', 'Unknown')
                        self.output.log(f"  {path[j]} --({relation})--> {path[j+1]}")
            
            return paths
        except nx.NetworkXNoPath:
            self.output.log("No paths found")
            return []
    
    def visualize_graph(self, layout_type='hierarchical', figsize=(20, 15), save_path=None):
        """
        Visualizes the graph with different colors for categories.
        
        Args:
            layout_type (str): Layout type ('spring', 'circular', 'hierarchical')
            figsize (tuple): Figure dimensions
            save_path (str): Path to save the image
        """
        if self.graph is None:
            self.output.log("Graph not available")
            return
        
        plt.figure(figsize=figsize)
        
        # Define the category colors
        categories = set(nx.get_node_attributes(self.graph, 'category').values())
        colors = cm.get_cmap('Set3')(np.linspace(0, 1, len(categories)))
        category_colors = dict(zip(categories, colors))

        # Choose the layout
        if layout_type == 'spring':
            pos = nx.spring_layout(self.graph, k=3, iterations=50)
        elif layout_type == 'circular':
            pos = nx.circular_layout(self.graph)
        elif layout_type == 'hierarchical':
            try:
                pos = nx.nx_agraph.graphviz_layout(self.graph, prog='dot')
            except:
                self.output.log("Layout gerarchico non disponibile, uso spring layout")
                pos = nx.spring_layout(self.graph)
        else:
            pos = nx.spring_layout(self.graph)

        # Draw the graph (simplified for compatibility)
        nx.draw_networkx_nodes(self.graph, pos, node_color='lightblue',
                              node_size=1000, alpha=0.8)

        # Draw the edges with different colors for each relation type
        relation_types = set(nx.get_edge_attributes(self.graph, 'relation_type').values())
        relation_colors = cm.get_cmap('tab10')(np.linspace(0, 1, len(relation_types)))
        relation_color_map = dict(zip(relation_types, relation_colors))
        
        for relation_type in relation_types:
            edges = [(u, v) for u, v, d in self.graph.edges(data=True) 
                    if d.get('relation_type') == relation_type]
            nx.draw_networkx_edges(self.graph, pos, edgelist=edges,
                                 edge_color=relation_color_map[relation_type],
                                 alpha=0.7, arrows=True, arrowsize=20,
                                 width=2)

        # Add labels to the nodes (abbreviated)
        labels = {node: node[:20] + '...' if len(node) > 20 else node
                 for node in self.graph.nodes()}
        nx.draw_networkx_labels(self.graph, pos, labels, font_size=8)

        # Create legend for categories
        legend_elements_cat = [Line2D([0], [0], marker='o', color='w',
                                         markerfacecolor=category_colors[cat], 
                                         markersize=10, label=cat) 
                              for cat in categories]
        
        legend1 = plt.legend(handles=legend_elements_cat, title="Threat Category",
                           loc='upper left', bbox_to_anchor=(1.05, 1))

        # Create legend for relation types
        legend_elements_rel = [Line2D([0], [0], color=relation_color_map[rel],
                                         linewidth=3, label=rel)
                              for rel in relation_types]

        legend2 = plt.legend(handles=legend_elements_rel, title="Relation Types",
                           loc='upper left', bbox_to_anchor=(1.05, 0.5))

        plt.gca().add_artist(legend1)  # Keep both legends

        plt.title("Attack Graph - Relationships between Space Cybersecurity Threats",
                 fontsize=16, fontweight='bold')
        plt.axis('off')
        plt.tight_layout()
        
        if save_path:
            plt.savefig(save_path, dpi=300, bbox_inches='tight')
            self.output.log(f"Graph saved to: {save_path}")

        plt.show()
    
    def create_category_network(self, figsize=(12, 8)):
        """Create a simplified graph of relationships between categories."""
        if self.df is None:
            return

        # Create a graph of categories
        cat_graph = nx.DiGraph()

        # Add relationships between categories with weights
        category_relations = self.df.groupby(['Source Category', 'Target Category']).size().reset_index(name='weight')
        
        for _, row in category_relations.iterrows():
            cat_graph.add_edge(row['Source Category'], row['Target Category'], 
                             weight=row['weight'])
        
        plt.figure(figsize=figsize)
        
        # Graph layout
        pos = nx.spring_layout(cat_graph, k=2, iterations=50)

        # Draw nodes
        nx.draw_networkx_nodes(cat_graph, pos, node_size=1000,
                              node_color='lightblue', alpha=0.8)

        # Draw edges
        nx.draw_networkx_edges(cat_graph, pos, width=2,
                              alpha=0.7, arrows=True, arrowsize=20)

        # Nodes labels
        nx.draw_networkx_labels(cat_graph, pos, font_size=10, font_weight='bold')

        # Edge labels (weights)
        edge_labels = {(u, v): str(cat_graph[u][v]['weight']) 
                      for u, v in cat_graph.edges()}
        nx.draw_networkx_edge_labels(cat_graph, pos, edge_labels, font_size=8)

        plt.title("Network of Threat Categories", fontsize=14, fontweight='bold')
        plt.axis('off')
        plt.tight_layout()
        plt.show()
    
    def export_to_gexf(self, output_path):
        """
        Export the graph to GEXF format for Gephi.

        Args:
            output_path (str): Output file path for the GEXF file.
        """
        if self.graph is None:
            self.output.log("Graph not available")
            return
        
        nx.write_gexf(self.graph, output_path)
        self.output.log(f"Graph exported to GEXF format: {output_path}")

    def run_interactive_analysis(self):
        """Run an interactive analysis where the user can choose specific threats using GUI."""
        if self.graph is None:
            messagebox.showerror("Error", "Graph not available for interactive analysis")
            return
        
        available_threats = list(self.graph.nodes())
        
        if not available_threats:
            messagebox.showerror("Error", "No threats available in the graph")
            return
        
        # Show initial info
        messagebox.showinfo("Interactive Analysis", 
                           f"Interactive Threat Analysis Mode\n\n"
                           f"Graph loaded with {len(available_threats)} threats\n\n"
                           f"You can choose specific threats to analyze in detail.")
        
        # Menu for analysis options using GUI
        def ask_analysis_option():
            """Ask user what type of analysis to perform"""
            root_temp = tk.Tk()
            root_temp.withdraw()
            
            class AnalysisOptionDialog:
                def __init__(self):
                    self.choice = None
                    self.root = tk.Toplevel()
                    self.root.title("🎯 Attack Graph Analysis Options")
                    self.root.geometry("700x500")
                    self.root.resizable(False, False)
                    self.root.configure(bg=COLORS['white'])
                    
                    # Center the window
                    self.root.transient()
                    self.root.grab_set()
                    
                    # Force window to front and keep on top
                    self.root.attributes('-topmost', True)
                    self.root.lift()
                    self.root.focus_force()
                    
                    # Remove topmost after 2 seconds to avoid annoying behavior
                    self.root.after(2000, lambda: self.root.attributes('-topmost', False))
                    
                    self.setup_ui()
                    
                def setup_ui(self):
                    # Header
                    header_frame = tk.Frame(self.root, bg=COLORS['primary'], height=80)
                    header_frame.pack(fill=tk.X)
                    header_frame.pack_propagate(False)
                    
                    title_label = tk.Label(header_frame, text="🎯 Attack Graph Analysis Options",
                                          font=('Segoe UI', 16, 'bold'),
                                          fg=COLORS['white'], bg=COLORS['primary'])
                    title_label.pack(expand=True)
                    
                    # Main content
                    content_frame = tk.Frame(self.root, bg=COLORS['white'])
                    content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
                    
                    # Info
                    info_label = tk.Label(content_frame, 
                                         text="Choose the type of analysis you want to perform:",
                                         font=('Segoe UI', 11), bg=COLORS['white'], fg=COLORS['dark'])
                    info_label.pack(pady=(0, 25))
                    
                    # Enhanced buttons with icons and descriptions
                    button_configs = [
                        {
                            'text': '📡 Threat Connections Analysis',
                            'desc': 'Analyze connections of a specific threat\n(predecessors and successors)',
                            'color': COLORS['primary'],
                            'hover': '#3a7ca8',
                            'choice': 1
                        },
                        {
                            'text': '🛤️ Path Analysis', 
                            'desc': 'Find attack paths between two threats\n(source → target)',
                            'color': COLORS['success'],
                            'hover': '#1e7e34',
                            'choice': 2
                        },
                        {
                            'text': '🔄 Combined Analysis',
                            'desc': 'Perform both threat connections and path analysis\n(comprehensive analysis)',
                            'color': COLORS['info'],
                            'hover': '#138496',
                            'choice': 3
                        },
                        {
                            'text': '❌ Exit Interactive Mode',
                            'desc': 'Return to main menu\n(close interactive analysis)',
                            'color': COLORS['danger'],
                            'hover': '#c82333',
                            'choice': 4
                        }
                    ]
                    
                    self.buttons = []
                    for config in button_configs:
                        # Button container for better layout
                        btn_container = tk.Frame(content_frame, bg=COLORS['white'])
                        btn_container.pack(fill=tk.X, pady=8)
                        
                        btn = tk.Button(btn_container,
                                       text=config['text'],
                                       font=('Segoe UI', 11, 'bold'),
                                       bg=config['color'], fg=COLORS['white'],
                                       relief='raised', bd=3,
                                       cursor='hand2',
                                       width=25, height=1,
                                       command=lambda c=config['choice']: self.set_choice(c))
                        btn.pack(side=tk.LEFT, padx=(0, 15))
                        
                        # Description label
                        desc_label = tk.Label(btn_container,
                                             text=config['desc'],
                                             font=('Segoe UI', 10),
                                             bg=COLORS['white'], fg=COLORS['gray'],
                                             justify=tk.LEFT)
                        desc_label.pack(side=tk.LEFT, expand=True, anchor='w')
                        
                        # Add hover effect
                        self.add_hover_effect(btn, config['hover'], config['color'])
                        self.buttons.append(btn)
                        
                def add_hover_effect(self, button, hover_color, normal_color):
                    def on_enter(e):
                        button.config(bg=hover_color)
                    def on_leave(e):
                        button.config(bg=normal_color)
                        
                    button.bind("<Enter>", on_enter)
                    button.bind("<Leave>", on_leave)
                    
                def set_choice(self, choice):
                    self.choice = choice
                    self.root.destroy()
            
            dialog = AnalysisOptionDialog()
            root_temp.wait_window(dialog.root)
            root_temp.destroy()
            
            return dialog.choice
        
        # Helper function for enhanced messageboxes
        def show_info_message(title, message, icon="info"):
            """Show an enhanced info message with better styling"""
            root_temp = tk.Tk()
            root_temp.withdraw()
            
            if icon == "success":
                messagebox.showinfo(f"✅ {title}", message)
            elif icon == "warning": 
                messagebox.showwarning(f"⚠️ {title}", message)
            elif icon == "error":
                messagebox.showerror(f"❌ {title}", message)
            else:
                messagebox.showinfo(f"💡 {title}", message)
                
            root_temp.destroy()
        
        def ask_yes_no(title, message):
            """Ask yes/no question with enhanced styling"""
            root_temp = tk.Tk()
            root_temp.withdraw()
            result = messagebox.askyesno(f"❓ {title}", message)
            root_temp.destroy()
            return result
        
        # Main interactive loop
        while True:
            choice = ask_analysis_option()
            
            if choice is None or choice == 4:
                messagebox.showinfo("Exit", "Exiting interactive mode")
                break
                
            elif choice == 1:
                # Threat connections analysis
                central_threat = interactive_threat_selection(available_threats, "central")
                
                if central_threat:
                    # Show analysis progress
                    messagebox.showinfo("Analysis Started", f"Analyzing connections for: {central_threat}")
                    
                    predecessors = list(self.graph.predecessors(central_threat))
                    successors = list(self.graph.successors(central_threat))
                    
                    # Save visualization
                    old_setting = THREAT_CONNECTION_ANALYSIS.get("save_visualization", False)
                    THREAT_CONNECTION_ANALYSIS["save_visualization"] = True
                    
                    result = self.analyze_threat_connections(central_threat)
                    
                    THREAT_CONNECTION_ANALYSIS["save_visualization"] = old_setting
                    
                    # Show results
                    if result:
                        show_info_message("Analysis Complete", 
                                           f"🎯 Threat connections analysis completed!\n\n"
                                           f"Central threat: {central_threat}\n"
                                           f"📥 Predecessors: {len(result['predecessors'])}\n"
                                           f"📤 Successors: {len(result['successors'])}\n\n"
                                           f"📊 Visualization saved to Output folder", "success")
                
            elif choice == 2:
                # Path analysis
                source_threat, target_threat = interactive_path_selection(available_threats)
                
                if source_threat and target_threat:
                    messagebox.showinfo("Analysis Started", f"Finding paths from:\n{source_threat}\nto:\n{target_threat}")
                    
                    paths = self.find_attack_paths(source_threat, target_threat, max_length=5)
                    
                    if paths:
                        # Ask if user wants to create visualization
                        create_viz = ask_yes_no("Paths Found", 
                                                        f"🎉 Found {len(paths)} path(s) between the selected threats!\n\n"
                                                        f"Source: {source_threat}\n"
                                                        f"Target: {target_threat}\n\n"
                                                        f"Would you like to create and save a path visualization?")
                        if create_viz:
                            self._create_combined_paths_graph(paths, source_threat, target_threat)
                            show_info_message("Visualization Complete", 
                                             f"🎨 Path visualization created and saved!\n\n"
                                             f"📁 Check the Output folder for the generated image", "success")
                        else:
                            show_info_message("Analysis Complete", 
                                             f"📊 Path analysis completed!\n\n"
                                             f"Found {len(paths)} path(s) between:\n"
                                             f"• {source_threat}\n"
                                             f"• {target_threat}", "success")
                    else:
                        show_info_message("No Paths Found", 
                                         f"🔍 No attack paths found between the selected threats.\n\n"
                                         f"Source: {source_threat}\n"
                                         f"Target: {target_threat}\n\n"
                                         f"💡 Try selecting different threats or check if they are connected.", "warning")
            
            elif choice == 3:
                # Both analyses
                messagebox.showinfo("Combined Analysis", "Running both analyses - threat connections and path analysis")
                
                # First: Threat connections
                central_threat = interactive_threat_selection(available_threats, "central")
                
                if central_threat:
                    old_setting = THREAT_CONNECTION_ANALYSIS.get("save_visualization", False)
                    THREAT_CONNECTION_ANALYSIS["save_visualization"] = True
                    result = self.analyze_threat_connections(central_threat)
                    THREAT_CONNECTION_ANALYSIS["save_visualization"] = old_setting
                    
                    messagebox.showinfo("Step 1 Complete", f"Threat connections completed for: {central_threat}")
                    
                    # Second: Path analysis
                    source_threat, target_threat = interactive_path_selection(available_threats)
                    
                    if source_threat and target_threat:
                        paths = self.find_attack_paths(source_threat, target_threat, max_length=5)
                        if paths:
                            self._create_combined_paths_graph(paths, source_threat, target_threat)
                            messagebox.showinfo("Analysis Complete", 
                                               f"Both analyses completed successfully!\n\n"
                                               f"Threat connections: {central_threat}\n"
                                               f"Path analysis: {source_threat} → {target_threat}\n"
                                               f"Paths found: {len(paths)}")
                        else:
                            messagebox.showwarning("Partial Success", 
                                                  f"Threat connections completed, but no paths found between selected threats")
            break
        
        show_info_message("Session Complete", 
                         "🎉 Interactive analysis session completed!\n\n"
                         "📁 All results and visualizations have been saved to the Output folder.\n"
                         "🔍 Check the generated files for detailed analysis results.", "success")

    def run_complete_analysis(self, interactive_mode=False):
        """Run a complete analysis and save everything to the output file.
        
        Args:
            interactive_mode (bool): If True, allows user to select specific threats interactively.
                                   If False, uses pre-configured automatic analysis.
        """
        from tkinter import messagebox
        
        if self.graph is None:
            messagebox.showerror("Error", "Graph not available for analysis")
            return
        
        available_threats = list(self.graph.nodes())
        
        if not available_threats:
            messagebox.showerror("Error", "No threats available in the graph")
            return
        
        self.output.log("🚀 STARTING COMPLETE ATTACK GRAPH ANALYSIS")
        
        if interactive_mode:
            self.output.log("🎮 INTERACTIVE MODE: User will select threats for analysis")
        else:
            self.output.log("🤖 AUTOMATIC MODE: Threats will be automatically determined")

        try:
            # Basic statistics
            self.get_graph_statistics()

            # Category analysis
            self.analyze_categories()

            # Centrality analysis
            self.analyze_centrality()

            # Critical paths analysis
            self.analyze_critical_paths()
            # Attack surface analysis
            self.analyze_attack_surface()

            # Specific threat connections analysis
            self.output.log("\n=== SPECIFIC THREAT NETWORK ANALYSIS ===")
            
            if interactive_mode:
                # User selects central threat for connections analysis
                messagebox.showinfo("Interactive Selection", 
                                   "Please select a central threat for connections analysis.")
                central_threat = interactive_threat_selection(available_threats, "central threat for connections analysis")
                if central_threat is None:
                    self.output.log("❌ User cancelled threat selection. Terminating analysis.")
                    return
                    
                # Save visualization for interactive mode
                old_setting = THREAT_CONNECTION_ANALYSIS.get("save_visualization", False)
                THREAT_CONNECTION_ANALYSIS["save_visualization"] = True
                self.analyze_threat_connections(central_threat)
                THREAT_CONNECTION_ANALYSIS["save_visualization"] = old_setting
            else:
                # Automatic selection using configured threat
                self.analyze_threat_connections()

            # Specific configurable paths analysis
            self.output.log("\n=== SPECIFIC CONFIGURABLE PATHS ANALYSIS ===")

            if interactive_mode:
                # User selects source and target threats for path analysis
                messagebox.showinfo("Interactive Selection", 
                                   "Please select source and target threats for path analysis.")
                source_threat, target_threat = interactive_path_selection(available_threats)
                if source_threat is None or target_threat is None:
                    self.output.log("❌ User cancelled path selection. Terminating analysis.")
                    return
                    
                max_len = SPECIFIC_PATH_ANALYSIS["max_path_length"]
                self.output.log(f"\n🎯 INTERACTIVE PATH: {source_threat} → {target_threat}")
                paths = self.find_attack_paths(source_threat, target_threat, max_len)
                
                # Create visualization for interactive paths
                if paths:
                    self._create_combined_paths_graph(paths, source_threat, target_threat)
                    
            else:
                # Automatic analysis using configured paths
                source = SPECIFIC_PATH_ANALYSIS["source_threat"]
                target = SPECIFIC_PATH_ANALYSIS["target_threat"]
                max_len = SPECIFIC_PATH_ANALYSIS["max_path_length"]

                self.output.log(f"\n🎯 MAIN PATH: {source} → {target}")
                paths = self.find_attack_paths(source, target, max_len)

                # Analyze multiple paths if configured
                if MULTIPLE_PATH_ANALYSIS:
                    self.output.log(f"\n🎯 MULTIPLE PATHS ANALYSIS ({len(MULTIPLE_PATH_ANALYSIS)} paths):")

                    for i, path_config in enumerate(MULTIPLE_PATH_ANALYSIS, 1):
                        self.output.log(f"\n--- PATH {i}: {path_config['description']} ---")
                        self.find_attack_paths(
                            path_config["source"], 
                            path_config["target"],
                            max_len
                        )
                
                # Create and save graphs for specific paths
                self.output.log("\n=== CREATION SPECIFIC PATH GRAPH ===")
                
                # Generate combined graphs for the main configured path
                if paths:
                    self._create_combined_paths_graph(paths, source, target)
                else:
                    self.output.log("No paths found for visualization")
            
            self.output.log("\n✅ ANALYSIS COMPLETED SUCCESSFULLY")
            
        except Exception as e:
            self.output.log(f"❌ Error occurred during analysis: {e}")
            self.output.log(traceback.format_exc())
        
        finally:
            self.output.close()

    def _create_combined_paths_graph(self, all_paths, source, target):
        """Create a combined graph with all found paths"""
        try:
            import matplotlib.pyplot as plt
            import networkx as nx

            # Create a graph that includes all nodes involved in the paths
            combined_graph = nx.DiGraph()
            # Add all nodes and edges from all paths
            for path in all_paths:
                for i in range(len(path) - 1):
                    source_node = path[i]
                    target_node = path[i + 1]
                    
                    if self.graph and self.graph.has_edge(source_node, target_node):
                        edge_data = self.graph[source_node][target_node]
                        combined_graph.add_edge(source_node, target_node, **edge_data)
                    else:
                        combined_graph.add_edge(source_node, target_node)

            # Figure configuration
            plt.figure(figsize=(20, 15))
            plt.suptitle(f'All Paths: {source} → {target}', 
                        fontsize=18, fontweight='bold')
            # Hierarchical layout: source at top, target at bottom
            pos = self._create_hierarchical_source_target_layout(combined_graph, source, target)

            # Node colors based on role
            node_colors = []
            for node in combined_graph.nodes():
                if node == source:
                    node_colors.append('#FF6B6B')  # Red for source
                elif node == target:
                    node_colors.append('#4ECDC4')  # Aqua green for target
                else:
                    node_colors.append('#FFD93D')  # Yellow for intermediate nodes

            # Draw the base graph
            nx.draw(combined_graph, pos,
                   node_color=node_colors,
                   node_size=2000,
                   with_labels=True,
                   labels={node: node.replace(' ', '\n') for node in combined_graph.nodes()},
                   font_size=6,
                   font_weight='bold',
                   arrows=True,
                   arrowsize=15,
                   arrowstyle='->',
                   edge_color='#BDC3C7',
                   width=1)

            # Highlight each path with different colors
            colors = ['#E74C3C', '#3498DB', '#9B59B6', '#E67E22', '#27AE60']
            
            for i, path in enumerate(all_paths):
                color = colors[i % len(colors)]

                # Draw the edges of the path
                path_edges = [(path[j], path[j + 1]) for j in range(len(path) - 1)]
                nx.draw_networkx_edges(combined_graph, pos,
                                     edgelist=path_edges,
                                     edge_color=color,
                                     width=3,
                                     arrows=True,
                                     arrowsize=20,
                                     arrowstyle='->')
            
                       
            # Path information - moved to top left
            paths_info = f"Percorsi trovati: {len(all_paths)}\n"
            for i, path in enumerate(all_paths, 1):
                paths_info += f"#{i}: {len(path)} nodi\n"
            
            plt.figtext(0.02, 0.98, paths_info, fontsize=10, verticalalignment='top',
                       bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray", alpha=0.9))

            # Legend
            legend_elements = [
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FF6B6B', 
                      markersize=15, label='Source'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#4ECDC4', 
                      markersize=15, label='Target'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FFD93D', 
                      markersize=15, label='Intermediate')
            ]
            
            for i, color in enumerate(colors[:len(all_paths)]):
                legend_elements.append(
                    Line2D([0], [0], color=color, linewidth=3, label=f'Path #{i+1}')
                )
            
            plt.legend(handles=legend_elements, loc='upper right')
            
            plt.tight_layout()
            # Save the combined graph
            source_clean = source.replace(' ', '_').replace(':', '').replace('/', '_')
            target_clean = target.replace(' ', '_').replace(':', '').replace('/', '_')
            filename = f"paths_combined_{source_clean}_{target_clean}.png"
            
            # Create Output directory if it doesn't exist
            output_dir = os.path.join(get_output_path(), "Output")
            os.makedirs(output_dir, exist_ok=True)
            filepath = os.path.join(output_dir, filename)
            
            plt.savefig(filepath, dpi=300, bbox_inches='tight', facecolor='white')
            plt.close()

            self.output.log(f"✅ Combined graph saved: {filepath}")

        except Exception as e:
            self.output.log(f"❌ Error creating combined graph: {e}")

    def _create_pseudo_hierarchical_layout(self, graph):
        """
        Create a pseudo-hierarchical layout using native NetworkX for specific paths

        Args:
            graph: The NetworkX graph to layout

        Returns:
            dict: A dictionary of node positions organized hierarchically
        """
        try:
            import networkx as nx

            # Calculate hierarchical levels based on topology
            # Find nodes without predecessors (roots)
            roots = [n for n in graph.nodes() if graph.in_degree(n) == 0]
            
            if not roots:
                # If there are no roots, use the node with the highest degree
                degrees = dict(graph.degree())
                if degrees:
                    roots = [max(degrees.keys(), key=lambda x: degrees[x])]
                else:
                    roots = list(graph.nodes())[:1]  # Take the first node if present

            # Calculate distances from each root
            levels = {}
            for node in graph.nodes():
                min_dist = float('inf')
                for root in roots:
                    try:
                        dist = nx.shortest_path_length(graph, root, node)
                        min_dist = min(min_dist, dist)
                    except nx.NetworkXNoPath:
                        continue
                
                if min_dist == float('inf'):
                    # Isolated node, place it at level 0
                    levels[node] = 0
                else:
                    levels[node] = min_dist

            # Organize nodes by level
            level_nodes = {}
            for node, level in levels.items():
                if level not in level_nodes:
                    level_nodes[level] = []
                level_nodes[level].append(node)
            
            # Hierarchical positions
            pos = {}
            max_level = max(level_nodes.keys()) if level_nodes else 0
            level_height = 3.0  # Space between levels

            for level, nodes in level_nodes.items():
                y = (max_level - level) * level_height  # Higher levels at the top

                # Distribute nodes horizontally within the level
                if len(nodes) == 1:
                    x_positions = [0]
                else:
                    x_positions = [(i - (len(nodes) - 1) / 2) * 2 for i in range(len(nodes))]
                
                for i, node in enumerate(nodes):
                    pos[node] = (x_positions[i], y)
            
            return pos
            
        except Exception as e:
            # Fallback: spring layout
            self.output.log(f"⚠️ Error with pseudo-hierarchical layout: {e}. Using spring layout")
            return nx.spring_layout(graph, k=3, iterations=50)

    def _create_hierarchical_source_target_layout(self, graph, source, target):
        """
        Create a hierarchical layout with source at top and target at bottom
        
        Args:
            graph: The NetworkX graph to layout
            source: Source node (placed at top)
            target: Target node (placed at bottom)
            
        Returns:
            dict: A dictionary of node positions
        """
        try:
            import networkx as nx
            
            pos = {}
            
            # Get all nodes from all paths
            all_nodes = set(graph.nodes())
            
            # Calculate distances from source
            try:
                distances_from_source = nx.single_source_shortest_path_length(graph, source)
            except:
                distances_from_source = {node: 0 for node in all_nodes}
            
            # Calculate distances to target (reverse graph)
            try:
                reverse_graph = graph.reverse()
                distances_to_target = nx.single_source_shortest_path_length(reverse_graph, target)
            except:
                distances_to_target = {node: 0 for node in all_nodes}
            
            # Organize nodes by level (distance from source)
            levels = {}
            max_level = 0
            for node in all_nodes:
                level = distances_from_source.get(node, 0)
                max_level = max(max_level, level)
                if level not in levels:
                    levels[level] = []
                levels[level].append(node)
            
            # Position nodes hierarchically
            level_height = 4.0  # Vertical spacing between levels
            node_spacing = 3.0  # Horizontal spacing between nodes
            
            for level, nodes in levels.items():
                # Y position: source at top (high Y), target at bottom (low Y)
                y = (max_level - level) * level_height
                
                # Sort nodes by distance to target for better visual flow
                nodes_sorted = sorted(nodes, key=lambda n: distances_to_target.get(n, 999))
                
                # X positions: center the nodes at each level
                num_nodes = len(nodes_sorted)
                if num_nodes == 1:
                    x_positions = [0]
                else:
                    total_width = (num_nodes - 1) * node_spacing
                    x_positions = [i * node_spacing - total_width/2 for i in range(num_nodes)]
                
                for i, node in enumerate(nodes_sorted):
                    pos[node] = (x_positions[i], y)
            
            # Force source at top and target at bottom
            if source in pos and target in pos:
                # Ensure source is at the highest Y
                max_y = max(pos[node][1] for node in pos)
                pos[source] = (pos[source][0], max_y + level_height)
                
                # Ensure target is at the lowest Y  
                min_y = min(pos[node][1] for node in pos if node != source)
                pos[target] = (pos[target][0], min_y - level_height)
            
            return pos
            
        except Exception as e:
            self.output.log(f"⚠️ Error with hierarchical source-target layout: {e}. Using spring layout")
            return nx.spring_layout(graph, k=3, iterations=50)

    def _create_hierarchical_threat_connections_layout(self, graph, central_threat, predecessors, successors):
        """
        Create a hierarchical layout for threat connections:
        - Central threat in the center
        - Predecessors on the left
        - Successors on the right
        
        Args:
            graph: The NetworkX graph to layout
            central_threat: The central threat node
            predecessors: List of predecessor nodes
            successors: List of successor nodes
            
        Returns:
            dict: A dictionary of node positions
        """
        try:
            import networkx as nx
            
            pos = {}
            
            # Constants for layout
            center_x = 0
            center_y = 0
            left_x_base = -8
            right_x_base = 8
            vertical_spacing = 3.0
            horizontal_spacing = 4.0
            
            # Place central threat at center
            if central_threat in graph.nodes():
                pos[central_threat] = (center_x, center_y)
            
            # Position predecessors on the left
            if predecessors:
                # Organize predecessors in levels based on distance from central threat
                left_levels = self._organize_nodes_by_distance(graph, central_threat, predecessors, reverse=True)
                
                for level, nodes in left_levels.items():
                    x_pos = left_x_base - (level * horizontal_spacing / 2)
                    
                    # Vertical positioning for nodes at same level
                    num_nodes = len(nodes)
                    if num_nodes == 1:
                        y_positions = [center_y]
                    else:
                        y_start = center_y + ((num_nodes - 1) * vertical_spacing) / 2
                        y_positions = [y_start - (i * vertical_spacing) for i in range(num_nodes)]
                    
                    for i, node in enumerate(nodes):
                        pos[node] = (x_pos, y_positions[i])
            
            # Position successors on the right
            if successors:
                # Organize successors in levels based on distance from central threat
                right_levels = self._organize_nodes_by_distance(graph, central_threat, successors, reverse=False)
                
                for level, nodes in right_levels.items():
                    x_pos = right_x_base + (level * horizontal_spacing / 2)
                    
                    # Vertical positioning for nodes at same level
                    num_nodes = len(nodes)
                    if num_nodes == 1:
                        y_positions = [center_y]
                    else:
                        y_start = center_y + ((num_nodes - 1) * vertical_spacing) / 2
                        y_positions = [y_start - (i * vertical_spacing) for i in range(num_nodes)]
                    
                    for i, node in enumerate(nodes):
                        # Add nodes that are both predecessors and successors twice
                        if node in predecessors:
                            # Create a duplicate node name for the right side
                            duplicate_node_name = f"{node}_successor_copy"
                            pos[duplicate_node_name] = (x_pos, y_positions[i])
                        else:
                            # Normal successor positioning
                            pos[node] = (x_pos, y_positions[i])
            
            # Add any remaining nodes not categorized
            remaining_nodes = set(graph.nodes()) - {central_threat} - set(predecessors) - set(successors)
            if remaining_nodes:
                # Place them at the bottom center
                y_bottom = center_y - 6
                num_remaining = len(remaining_nodes)
                if num_remaining == 1:
                    x_positions = [center_x]
                else:
                    x_start = center_x - ((num_remaining - 1) * 2.0) / 2
                    x_positions = [x_start + (i * 2.0) for i in range(num_remaining)]
                
                for i, node in enumerate(remaining_nodes):
                    if node not in pos:
                        pos[node] = (x_positions[i], y_bottom)
            
            return pos
            
        except Exception as e:
            self.output.log(f"⚠️ Error with hierarchical threat connections layout: {e}. Using spring layout")
            return nx.spring_layout(graph, k=3, iterations=50)

    def _organize_nodes_by_distance(self, graph, central_node, nodes, reverse=False):
        """
        Organize nodes by their distance from the central node
        
        Args:
            graph: NetworkX graph
            central_node: Central reference node
            nodes: List of nodes to organize
            reverse: If True, use reverse graph (for predecessors)
            
        Returns:
            dict: Dictionary with levels as keys and lists of nodes as values
        """
        try:
            import networkx as nx
            
            if reverse:
                work_graph = graph.reverse()
            else:
                work_graph = graph
            
            levels = {}
            
            for node in nodes:
                try:
                    if reverse:
                        # For predecessors, calculate distance from node to central
                        distance = nx.shortest_path_length(work_graph, central_node, node)
                    else:
                        # For successors, calculate distance from central to node
                        distance = nx.shortest_path_length(work_graph, central_node, node)
                except nx.NetworkXNoPath:
                    distance = 1  # Default distance if no path
                
                if distance not in levels:
                    levels[distance] = []
                levels[distance].append(node)
            
            return levels
            
        except Exception as e:
            # Fallback: all nodes at level 1
            return {1: list(nodes)}
    

def modify_configuration():
    """
    Helper function to easily modify the configuration.
    Modify the global variables above instead of using this function.
    """
    ##print("💡 To modify the configuration, change the variables at the beginning of the file:")
    ##print("   - THREAT_FILE_NAME: for the CSV file of threats to analyze")
    ##print("   - SPECIFIC_PATH_ANALYSIS: for the main path")
    ##print("   - MULTIPLE_PATH_ANALYSIS: for additional paths")
    ##print("   - ANALYSIS_PARAMETERS: for the analysis parameters")


def select_csv_file():
    """
    Opens a file dialog to select the CSV file with threats to analyze.
    
    Returns:
        str: Path to the selected CSV file, or None if cancelled
    """
    # Create a hidden root window
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    
    # Configure the file dialog
    file_types = [
        ("CSV files", "*.csv"),
        ("All files", "*.*")
    ]
    
    # Show file dialog
    try:
        file_path = filedialog.askopenfilename(
            title="Select CSV file with threats to analyze",
            filetypes=file_types,
            initialdir="."
        )
        
        # Check if user cancelled the dialog
        if not file_path:  # Empty string or None
            root.destroy()
            return None
        
        if file_path:
            # Verify the file exists and is readable
            if not os.path.exists(file_path):
                messagebox.showerror("Error", f"File {file_path} not found!")
                root.destroy()
                return None
            
            # Try to read the file to validate it's a valid CSV
            try:
                df = pd.read_csv(file_path, sep=';')
                if 'THREAT' not in df.columns:
                    messagebox.showerror(
                        "Invalid File",
                        f"The selected file must contain a 'THREAT' column.\n\n"
                        f"Columns found: {list(df.columns)}\n\n"
                        f"Expected format:\n"
                        f"THREAT;Likelihood;Impact;Risk"
                    )
                    root.destroy()
                    return None
                
                # Show confirmation with file info
                num_threats = len(df)
                messagebox.showinfo(
                    "File Selected",
                    f"Selected file: {os.path.basename(file_path)}\n"
                    f"Number of threats: {num_threats}\n"
                    f"Columns: {list(df.columns)}"
                )
                
            except Exception as e:
                messagebox.showerror(
                    "File Reading Error",
                    f"Error reading the selected file:\n{str(e)}\n\n"
                    f"Please ensure the file is a valid CSV with ';' separator."
                )
                root.destroy()
                return None
        
        root.destroy()
        return file_path
        
    except Exception as e:
        messagebox.showerror("Error", f"Error selecting file: {str(e)}")
        root.destroy()
        return None

def main():
    """Main function to test the analyzer."""
    
    # Show file selection dialog
    #print("🔍 SELECT CSV FILE WITH THREATS TO ANALYZE")
    #print("=" * 50)
    
    selected_file = select_csv_file()
    
    if selected_file is None:
        # User cancelled file selection - show message and exit
        try:
            messagebox.showinfo("Cancelled", "File selection cancelled. Exiting application.")
        except:
            pass  # In case tkinter window was already destroyed
        return
    
    #print(f"✅ Selected file: {selected_file}")
    
    # Update the global variable with the selected file
    global THREAT_FILE_NAME
    THREAT_FILE_NAME = selected_file
    
    # Show the current configuration
    #print_configuration()
    # File paths
    csv_path = os.path.join(get_base_path(), "attack_graph_threat_relations.csv")
    # Use the selected file for threats - ensure it has the correct base path
    if os.path.isabs(THREAT_FILE_NAME):
        subset_path = THREAT_FILE_NAME
    else:
        subset_path = os.path.join(get_base_path(), THREAT_FILE_NAME)

    # Check if the files exist
    #if not os.path.exists(csv_path):
        ##print(f"❌ File {csv_path} not found!")
        ##print("📁 Available files in the directory:")
    #    for file in os.listdir("."):
    #        if file.endswith(".csv"):
                ##print(f"   - {file}")
    #    return

    ##print(f"📊 Using relationships from file: {csv_path}")

    #if os.path.exists(subset_path):
        ##print(f"🎯 Filtering threats from file: {subset_path}")
    #else:
        ##print(f"⚠️  Subset file '{subset_path}' not found. Analyzing all threats.")

    # Create the analyzer with subset and output to file
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    analyzer = AttackGraphAnalyzer(
        csv_file_path=csv_path, 
        subset_file_path=subset_path,
        output_file=f"attack_graph_analysis_{timestamp}.txt"
    )

    # Ask user for analysis mode using GUI
    def ask_analysis_mode():
        """Ask user to choose analysis mode using an enhanced GUI dialog"""
        root = tk.Tk()
        root.withdraw()  # Hide the root window
        
        class ModeSelectionDialog:
            def __init__(self):
                self.choice = None
                self.root = tk.Toplevel()
                self.root.title("🚀 CRAAL Space Threat Analyzer")
                self.root.geometry("600x500")
                self.root.resizable(True, True)
                
                # Center the window
                self.root.transient()
                self.root.grab_set()
                
                # Force window to front and keep on top
                self.root.attributes('-topmost', True)
                self.root.lift()
                self.root.focus_force()
                
                # Remove topmost after 2 seconds to avoid annoying behavior
                self.root.after(2000, lambda: self.root.attributes('-topmost', False))
                
                # Set window icon and style
                try:
                    self.root.iconbitmap()  # Use default
                except:
                    pass
                
                self.setup_ui()
                
            def setup_ui(self):
                # Main frame with gradient-like effect
                main_frame = tk.Frame(self.root, bg=COLORS['light'], relief='raised', bd=2)
                main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
                
                # Header with title and icon
                header_frame = tk.Frame(main_frame, bg=COLORS['primary'], height=80)
                header_frame.pack(fill=tk.X, padx=5, pady=5)
                header_frame.pack_propagate(False)
                
                title_label = tk.Label(header_frame, 
                                      text="🛰️ CRAAL Space Threat Analyzer", 
                                      font=('Segoe UI', 16, 'bold'),
                                      fg=COLORS['white'], bg=COLORS['primary'])
                title_label.pack(expand=True)
                
                subtitle_label = tk.Label(header_frame, 
                                         text="Cybersecurity Risk Assessment & Attack Learning", 
                                         font=('Segoe UI', 11, 'italic'),
                                         fg=COLORS['light'], bg=COLORS['primary'])
                subtitle_label.pack()
                
                # Content frame
                content_frame = tk.Frame(main_frame, bg=COLORS['light'])
                content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
                
                # Instructions
                instruction_label = tk.Label(content_frame, 
                                            text="Choose the analysis mode:",
                                            font=('Segoe UI', 11, 'bold'),
                                            bg=COLORS['light'], fg=COLORS['dark'])
                instruction_label.pack(pady=(0, 20))
                
                # Buttons frame
                buttons_frame = tk.Frame(content_frame, bg=COLORS['light'])
                buttons_frame.pack(pady=20)
                
                # Interactive mode button
                interactive_btn = tk.Button(buttons_frame, 
                                           text="🎮 Interactive Analysis\nSelect specific threats with GUI",
                                           font=('Segoe UI', 11, 'bold'),
                                           bg=COLORS['success'], fg=COLORS['white'],
                                           relief='raised', bd=3,
                                           width=33, height=2,
                                           cursor='hand2',
                                           command=lambda: self.set_choice(True))
                interactive_btn.pack(pady=10)
                
                # Auto mode button  
                auto_btn = tk.Button(buttons_frame, 
                                    text="🤖 Automatic Analysis\nComplete analysis with preset configuration",
                                    font=('Segoe UI', 11, 'bold'),
                                    bg=COLORS['primary'], fg=COLORS['white'],
                                    relief='raised', bd=3,
                                    width=33, height=2,
                                    cursor='hand2',
                                    command=lambda: self.set_choice(False))
                auto_btn.pack(pady=10)
                
                # Cancel button
                cancel_btn = tk.Button(buttons_frame, 
                                      text="❌ Cancel",
                                      font=('Segoe UI', 11),
                                      bg=COLORS['danger'], fg=COLORS['white'],
                                      relief='raised', bd=3,
                                      width=15, height=1,
                                      cursor='hand2',
                                      command=lambda: self.set_choice(None))
                cancel_btn.pack(pady=10)
                
                # Footer
                #footer_frame = tk.Frame(main_frame, bg='#e5e7eb', height=40)
                #footer_frame.pack(fill=tk.X, side=tk.BOTTOM, padx=5, pady=5)
                #footer_frame.pack_propagate(False)
                
                #footer_label = tk.Label(footer_frame, 
                #                       text="💡 Interactive mode allows you to manually select specific threats for detailed analysis",
                #                       font=('Segoe UI', 9),
                #                       bg='#e5e7eb', fg='#6b7280')
                #footer_label.pack(expand=True)
                
                # Bind hover effects
                self.add_hover_effects(interactive_btn, '#059669', '#10b981')
                self.add_hover_effects(auto_btn, '#2563eb', '#3b82f6')
                self.add_hover_effects(cancel_btn, '#dc2626', '#ef4444')
                
            def add_hover_effects(self, button, hover_color, normal_color):
                def on_enter(e):
                    button.config(bg=hover_color)
                def on_leave(e):
                    button.config(bg=normal_color)
                    
                button.bind("<Enter>", on_enter)
                button.bind("<Leave>", on_leave)
                
            def set_choice(self, choice):
                self.choice = choice
                self.root.destroy()
        
        dialog = ModeSelectionDialog()
        root.wait_window(dialog.root)
        root.destroy()
        
        return dialog.choice
    
    mode_choice = ask_analysis_mode()
    
    if mode_choice is None:  # User clicked Cancel
        messagebox.showinfo("Cancelled", "Analysis cancelled by user")
        return
    elif mode_choice:  # User clicked Yes (Interactive)
        messagebox.showinfo("Interactive Mode", "Starting interactive analysis with GUI threat selection...")
        analyzer.run_complete_analysis(interactive_mode=True)
    else:  # User clicked No (Automatic)
        messagebox.showinfo("Automatic Mode", "Starting automatic analysis with pre-configured settings...")
        analyzer.run_complete_analysis(interactive_mode=False)

    # Generate visualizations (if matplotlib works)
    try:
        ##print("\n=== GENERATING THE DISPLAY ===")
        analyzer.create_category_network()
        
        # Create Output directory if it doesn't exist
        output_dir = os.path.join(get_output_path(), "Output")
        os.makedirs(output_dir, exist_ok=True)
        
        analyzer.visualize_graph(layout_type='spring', save_path=os.path.join(output_dir, 'attack_graph.png'))

        # Export for Gephi
        analyzer.export_to_gexf(os.path.join(output_dir, 'attack_graph.gexf'))
        
    except Exception as e:
        return
        ##print(f"⚠️  Error in visualizations: {e}")
        ##print("Textual analysis has been completed and saved to the file.")


def interactive_threat_selection(graph_nodes, selection_type="threat"):
    """
    Allows user to interactively select a threat using a GUI dialog.
    
    Args:
        graph_nodes (list): List of available threat nodes
        selection_type (str): Type of selection ("threat", "source", "target")
        
    Returns:
        str: Selected threat name or None if cancelled
    """
    import tkinter as tk
    from tkinter import ttk, messagebox
    import random

    if not graph_nodes:
        messagebox.showerror("Error", f"No threats available for {selection_type} selection")
        return None
    
    
    # Sort threats alphabetically for easier browsing
    sorted_threats = sorted(graph_nodes)
    
    class ThreatSelectorDialog:
        def __init__(self, threats, selection_type):
            self.threats = threats
            self.selection_type = selection_type
            self.selected_threat = None
            self.filtered_threats = threats.copy()
            
            # Create main window with enhanced styling
            self.root = tk.Toplevel()
            self.root.title(f"🎯 Select {selection_type.capitalize()} Threat")
            self.root.geometry("700x700")
            self.root.resizable(True, True)
            self.root.configure(bg=COLORS['white'])
            
            # Center the window
            self.root.transient()
            self.root.grab_set()
            
            # Force window to front and keep on top
            self.root.attributes('-topmost', True)
            self.root.lift()
            self.root.focus_force()
            
            # Remove topmost after 2 seconds to avoid annoying behavior
            self.root.after(2000, lambda: self.root.attributes('-topmost', False))
            
            self.setup_ui()
            
        def setup_ui(self):
            # Header frame with gradient-like effect
            header_frame = tk.Frame(self.root, bg=COLORS['primary'], height=80)
            header_frame.pack(fill=tk.X, padx=0, pady=0)
            header_frame.pack_propagate(False)
            
            # Header title with icon
            icon_dict = {'source': '🚀', 'target': '🎯', 'central': '⭐', 'threat': '🔍'}
            icon = icon_dict.get(self.selection_type, '🔍')
            
            title_label = tk.Label(header_frame, 
                                  text=f"{icon} Select {self.selection_type.capitalize()} Threat",
                                  font=('Segoe UI', 16, 'bold'),
                                  fg=COLORS['white'], bg=COLORS['primary'])
            title_label.pack(expand=True)
            
            # Main content frame
            content_frame = tk.Frame(self.root, bg=COLORS['white'])
            content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
            
            # Info panel
            info_frame = tk.Frame(content_frame, bg=COLORS['light'], relief='ridge', bd=2)
            info_frame.pack(fill=tk.X, pady=(0, 15))
            
            info_text = f"📊 Available threats: {len(self.threats)}   |   💡 Use search to filter   |   🎲 Random selection available"
            info_label = tk.Label(info_frame, text=info_text,
                                 font=('Segoe UI', 11), bg=COLORS['light'], fg=COLORS['dark'],
                                 pady=8)
            info_label.pack()
            
            # Search frame with enhanced styling
            search_frame = tk.LabelFrame(content_frame, text="🔍 Search & Filter", 
                                        font=('Segoe UI', 11, 'bold'),
                                        bg=COLORS['white'], fg=COLORS['primary'],
                                        relief='groove', bd=2)
            search_frame.pack(fill=tk.X, pady=(0, 15))
            
            search_inner = tk.Frame(search_frame, bg=COLORS['white'])
            search_inner.pack(fill=tk.X, padx=10, pady=10)
            
            tk.Label(search_inner, text="Search:", font=('Segoe UI', 11, 'bold'),
                    bg=COLORS['white'], fg=COLORS['dark']).pack(side=tk.LEFT, padx=(0, 8))
            
            self.search_var = tk.StringVar()
            search_entry = tk.Entry(search_inner, textvariable=self.search_var,
                                   font=('Segoe UI', 11), relief='solid', bd=1,
                                   highlightthickness=2, highlightcolor=COLORS['primary'])
            search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
            search_entry.bind('<KeyRelease>', self.filter_threats)
            
            clear_btn = tk.Button(search_inner, text="Clear", 
                                 font=('Segoe UI', 11), bg=COLORS['gray'], fg=COLORS['white'],
                                 relief='raised', bd=2, cursor='hand2',
                                 command=self.clear_search)
            clear_btn.pack(side=tk.RIGHT)
            
            # Main selection frame
            selection_frame = tk.LabelFrame(content_frame, text="🎯 Threat Selection",
                                           font=('Segoe UI', 11, 'bold'),
                                           bg=COLORS['white'], fg=COLORS['primary'],
                                           relief='groove', bd=2)
            selection_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
            
            # Listbox frame
            list_container = tk.Frame(selection_frame, bg='#f8fafc')
            list_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            # Scrollbars with custom styling
            v_scrollbar = tk.Scrollbar(list_container, orient=tk.VERTICAL)
            h_scrollbar = tk.Scrollbar(list_container, orient=tk.HORIZONTAL)
            
            # Enhanced listbox
            self.listbox = tk.Listbox(list_container,
                                     yscrollcommand=v_scrollbar.set,
                                     xscrollcommand=h_scrollbar.set,
                                     font=('Consolas', 10),
                                     selectbackground='#3b82f6',
                                     selectforeground='white',
                                     activestyle='dotbox',
                                     relief='solid', bd=1)
            
            # Configure scrollbars
            v_scrollbar.config(command=self.listbox.yview)
            h_scrollbar.config(command=self.listbox.xview)
            
            # Grid layout for listbox and scrollbars
            self.listbox.grid(row=0, column=0, sticky="news")
            v_scrollbar.grid(row=0, column=1, sticky="ns")
            h_scrollbar.grid(row=1, column=0, sticky="we")
            
            list_container.grid_columnconfigure(0, weight=1)
            list_container.grid_rowconfigure(0, weight=1)
            
            # Double-click to select
            self.listbox.bind('<Double-Button-1>', self.on_double_click)
            
            # Enhanced buttons frame
            button_frame = tk.Frame(content_frame, bg='#f8fafc')
            button_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Button styling
            button_style = {
                'font': ('Segoe UI', 11, 'bold'),
                'relief': 'raised',
                'bd': 3,
                'cursor': 'hand2',
                'width': 15,
                'height': 1
            }
            
            random_btn = tk.Button(button_frame, text="🎲 Random", 
                                  bg='#8b5cf6', fg='white',
                                  command=self.select_random, **button_style)
            random_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            select_btn = tk.Button(button_frame, text="✅ Select", 
                                  bg='#10b981', fg='white',
                                  command=self.select_current, **button_style)
            select_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            skip_btn = tk.Button(button_frame, text="⏭️ Skip", 
                                bg='#f59e0b', fg='white',
                                command=self.skip_selection, **button_style)
            skip_btn.pack(side=tk.LEFT, padx=(0, 8))
            
            cancel_btn = tk.Button(button_frame, text="❌ Cancel", 
                                  bg='#ef4444', fg='white',
                                  command=self.cancel, **button_style)
            cancel_btn.pack(side=tk.RIGHT)
            
            # Add hover effects
            self.add_hover_effects(random_btn, '#7c3aed', '#8b5cf6')
            self.add_hover_effects(select_btn, '#059669', '#10b981')
            self.add_hover_effects(skip_btn, '#d97706', '#f59e0b')
            self.add_hover_effects(cancel_btn, '#dc2626', '#ef4444')
            self.add_hover_effects(clear_btn, '#4b5563', '#6b7280')
            
            # Populate the list initially
            self.update_listbox()
            
            # Set focus to search entry
            search_entry.focus()
            
        def add_hover_effects(self, button, hover_color, normal_color):
            def on_enter(e):
                button.config(bg=hover_color)
            def on_leave(e):
                button.config(bg=normal_color)
                
            button.bind("<Enter>", on_enter)
            button.bind("<Leave>", on_leave)
            
        def filter_threats(self, event=None):
            search_term = self.search_var.get().lower()
            if search_term:
                self.filtered_threats = [t for t in self.threats if search_term in t.lower()]
            else:
                self.filtered_threats = self.threats.copy()
            self.update_listbox()
            
        def clear_search(self):
            self.search_var.set("")
            self.filtered_threats = self.threats.copy()
            self.update_listbox()
            
        def update_listbox(self):
            self.listbox.delete(0, tk.END)
            for threat in self.filtered_threats:
                self.listbox.insert(tk.END, threat)
                
        def on_double_click(self, event=None):
            self.select_current()
            
        def select_current(self):
            selection = self.listbox.curselection()
            if selection:
                self.selected_threat = self.filtered_threats[selection[0]]
                self.root.destroy()
            else:
                messagebox.showwarning("No Selection", "Please select a threat from the list")
                
        def select_random(self):
            if self.filtered_threats:
                self.selected_threat = random.choice(self.filtered_threats)
                self.root.destroy()
            else:
                messagebox.showwarning("No Threats", "No threats available for random selection")
                
        def skip_selection(self):
            self.selected_threat = None
            self.root.destroy()
            
        def cancel(self):
            self.selected_threat = None
            self.root.destroy()
    
    # Create dialog
    dialog = ThreatSelectorDialog(sorted_threats, selection_type)
    dialog.root.wait_window()
    
    return dialog.selected_threat


def interactive_path_selection(graph_nodes):
    """
    Allows user to interactively select source and target threats using GUI dialogs.
    
    Args:
        graph_nodes (list): List of available threat nodes
        
    Returns:
        tuple: (source_threat, target_threat) or (None, None) if cancelled
    """
    # First, select source threat
    source_threat = interactive_threat_selection(graph_nodes, "source")

    if source_threat is None:
        return None, None
    
    # Select target (excluding the source)
    available_targets = [node for node in graph_nodes if node != source_threat]
    if not available_targets:
        messagebox.showerror("Error", f"No target threats available after excluding source '{source_threat}'")
        return source_threat, None
    
    # Show confirmation of source selection
    result = messagebox.askquestion("Source Selected", 
                                   f"Source threat selected: {source_threat}\n\nProceed to select target threat?",
                                   icon='question')
    if result == 'no':
        return None, None
    
    # Select target threat
    target_threat = interactive_threat_selection(available_targets, "target")
    
    if target_threat is None:
        return source_threat, None
    
    # Show final confirmation
    messagebox.showinfo("Path Analysis Configured", 
                       f"Path Analysis Setup Complete:\n\nSource: {source_threat}\nTarget: {target_threat}")
    
    return source_threat, target_threat


if __name__ == "__main__":
    main()
