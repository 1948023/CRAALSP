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
        # PyInstaller stores data files in sys._MEIPASS
        return getattr(sys, '_MEIPASS', os.path.dirname(sys.executable))
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
    ##print("⚠️  Scipy not available. Some centrality features will be limited.")

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
            ##print(f"⚠️  Unable to create output file: {e}")
            self.file_handle = None
    
    def log(self, message):
        """Writes a message both to console and file."""
        ##print(message)
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
            
            if len(subgraph.nodes()) == 0:
                self.output.log("   ⚠️ No nodes to visualize")
                return
            
            # Configure the visualization
            plt.figure(figsize=(16, 12))
            plt.clf()

            # Layout - use spring layout with the central threat fixed in the center
            pos = nx.spring_layout(subgraph, k=3, iterations=50,
                                 fixed=[target_threat] if target_threat in subgraph.nodes() else None,
                                 pos={target_threat: (0, 0)} if target_threat in subgraph.nodes() else None)

            # Colors for different types of nodes
            node_colors = []
            node_sizes = []
            
            for node in subgraph.nodes():
                if node == target_threat:
                    node_colors.append('#FF4444')  # Red for the central threat
                    node_sizes.append(2000)
                elif node in successors:
                    node_colors.append('#44FF44')  # Green for successors
                    node_sizes.append(1200)
                elif node in predecessors:
                    node_colors.append('#4444FF')  # Blue for predecessors
                    node_sizes.append(1200)
                else:
                    node_colors.append('#FFAA44')  # Orange for second-level neighbors
                    node_sizes.append(800)

            # Draw the nodes
            nx.draw_networkx_nodes(subgraph, pos,
                                 node_color=node_colors,  # type: ignore
                                 node_size=node_sizes,  # type: ignore
                                 alpha=0.8,
                                 edgecolors='black',
                                 linewidths=2)

            # Draw edges with different colors
            edge_colors = []
            edge_styles = []
            
            for edge in subgraph.edges():
                source, target = edge
                if source == target_threat or target == target_threat:
                    edge_colors.append('#333333')  # Dark gray for direct connections
                    edge_styles.append('solid')
                else:
                    edge_colors.append('#888888')  # Light gray for indirect connections
                    edge_styles.append('dashed')
            
            nx.draw_networkx_edges(subgraph, pos,
                                 edge_color=edge_colors,  # type: ignore
                                 style=edge_styles,  # type: ignore
                                 width=2,
                                 alpha=0.7,
                                 arrows=True,
                                 arrowsize=20,
                                 arrowstyle='->')
            
            # Node Label with abbreviations
            labels = {}
            for node in subgraph.nodes():
                # Abbreviate long names
                #if len(node) > 30:
                #    label = node[:27] + "..."
                #else:
                #    label = node
                label = node
                labels[node] = label
            
            nx.draw_networkx_labels(subgraph, pos, labels,
                                  font_size=9,
                                  font_weight='bold',
                                  font_color='white',
                                  bbox=dict(boxstyle='round,pad=0.3', 
                                          facecolor='black', 
                                          alpha=0.7))
            
            # Edge Label if configured
            if THREAT_CONNECTION_ANALYSIS.get("show_relation_types", True):
                edge_labels = {}
                for edge in subgraph.edges():
                    edge_data = subgraph[edge[0]][edge[1]]
                    relation_type = edge_data.get('relation_type', 'Unknown')
                    # Abbreviate long relation types
                    #if len(relation_type) > 15:
                    #    relation_type = relation_type[:12] + "..."
                    edge_labels[edge] = relation_type
                
                nx.draw_networkx_edge_labels(subgraph, pos, edge_labels,
                                           font_size=7,
                                           font_color='darkred',
                                           bbox=dict(boxstyle='round,pad=0.2',
                                                   facecolor='white',
                                                   alpha=0.8))

            # Title and legend
            plt.title(f"Threat Connections: {target_threat}", 
                     fontsize=16, fontweight='bold', pad=20)

            # Create legend
            legend_elements = [
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FF4444', 
                          markersize=15, label='Central Threat  '),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#4444FF', 
                          markersize=12, label='Predecessors (enable)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#44FF44', 
                          markersize=12, label='Successors (enabled)'),
                Line2D([0], [0], marker='o', color='w', markerfacecolor='#FFAA44', 
                          markersize=10, label='Second Level'),
                Line2D([0], [0], color='#333333', linewidth=2, label='Direct Connection'),
                Line2D([0], [0], color='#888888', linewidth=2, linestyle='--', label='Indirect Connection')
            ]
            
            plt.legend(handles=legend_elements, loc='upper right', bbox_to_anchor=(1.15, 1))
            
            # Additional info
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

    def run_complete_analysis(self):
        """Run a complete analysis and save everything to the output file."""
        self.output.log("🚀 STARTING COMPLETE ATTACK GRAPH ANALYSIS")

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
            self.analyze_threat_connections()

            # Specific configurable paths analysis
            self.output.log("\n=== SPECIFIC CONFIGURABLE PATHS ANALYSIS ===")

            # Analyze the main configured path
            source = SPECIFIC_PATH_ANALYSIS["source_threat"]
            target = SPECIFIC_PATH_ANALYSIS["target_threat"]
            max_len = SPECIFIC_PATH_ANALYSIS["max_path_length"]

            self.output.log(f"\n🎯 MAIN PATH: {source} → {target}")
            self.find_attack_paths(source, target, max_len)

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
            source = SPECIFIC_PATH_ANALYSIS["source_threat"]
            target = SPECIFIC_PATH_ANALYSIS["target_threat"]
            main_paths = self.find_attack_paths(source, target, max_len)
            
            if main_paths:
                self._create_combined_paths_graph(main_paths, source, target)
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
            # Pseudo-hierarchical layout for the combined graph
            pos = self._create_pseudo_hierarchical_layout(combined_graph)

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
            
                       
            # Path information
            paths_info = f"Percorsi trovati: {len(all_paths)}\n"
            for i, path in enumerate(all_paths, 1):
                paths_info += f"#{i}: {len(path)} nodi\n"
            
            plt.figtext(0.02, 0.02, paths_info, fontsize=10,
                       bbox=dict(boxstyle="round,pad=0.3", facecolor="lightgray"))

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
    

#def print_configuration():
#    """Print the current configuration for verification."""
    ##print("\n" + "="*70)
    ##print("🔧 CONFIGURATION ATTACK GRAPH ANALYZER")
    ##print("="*70)

    ##print(f"\n📂 THREAT FILE CONFIGURATED:")
    ##print(f"   {THREAT_FILE_NAME}")

    ##print(f"\n📍 MAIN PATH:")
    ##print(f"   From: {SPECIFIC_PATH_ANALYSIS['source_threat']}")
    ##print(f"   To:  {SPECIFIC_PATH_ANALYSIS['target_threat']}")
    ##print(f"   Max length: {SPECIFIC_PATH_ANALYSIS['max_path_length']}")

#    if MULTIPLE_PATH_ANALYSIS:
        ##print(f"\n📍 MULTIPLE PATHS ({len(MULTIPLE_PATH_ANALYSIS)}):")
#        for i, path in enumerate(MULTIPLE_PATH_ANALYSIS, 1):
            ##print(f"   {i}. {path['description']}")
            ##print(f"      {path['source']} → {path['target']}")

    ##print(f"\n⚙️ ANALYSIS PARAMETERS:")
#    for key, value in ANALYSIS_PARAMETERS.items():
        ##print(f"   {key}: {value}")
    
    ##print("="*70)


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
        #print("❌ No file selected. Exiting.")
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

    # Run complete analysis
    analyzer.run_complete_analysis()

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


if __name__ == "__main__":
    main()
