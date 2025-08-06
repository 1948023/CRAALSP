# Cyber Risk Assessment Across Lifecycle of Space Program (CRAALSP)

A comprehensive cybersecurity risk assessment toolkit designed specifically for space missions and satellite systems. This suite provides tools for threat analysis, risk evaluation, and attack graph visualization to support space program security assessments.

## 🚀 Overview

This tool suite was developed as part of a thesis work for space program risk assessment and consists of four main components with both Python scripts and standalone executables:

1. **BID Phase Tool** - Initial risk assessment based on project category
2. **Risk Assessment 0-A** - Preliminary threat and vulnerability analysis  
3. **Risk Assessment** - Comprehensive risk evaluation with detailed criteria
4. **Attack Graph Analyzer** - Advanced threat relationship analysis and visualization with network graphs

### 🎯 Key Features

- **Dual Execution Modes**: Python scripts and standalone Windows executables (.exe)
- **Main Launcher Interface**: Unified GUI to run all tools
- **Professional Reporting**: Automated Word document generation
- **Advanced Visualizations**: Network graphs, attack paths, and statistical analysis
- **Interactive Analysis**: Real-time threat selection and path exploration
- **Export Capabilities**: Multiple formats (Word, CSV, PNG, GEXF)

## 📋 Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Standalone Executables](#standalone-executables)
- [Tool Descriptions](#tool-descriptions)
- [Attack Graph Analyzer - Advanced Features](#attack-graph-analyzer---advanced-features)
- [File Structure](#file-structure)
- [Usage Examples](#usage-examples)
- [Configuration](#configuration)
- [Dependencies](#dependencies)
- [Data Formats](#data-formats)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)

## 🛠️ Installation

### Prerequisites

- Python 3.7 or higher (for Python scripts)
- Windows 10/11 (for standalone executables)
- No additional software required for .exe files

### Required Libraries (Python Scripts Only)

```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install tkinter
pip install python-docx
pip install pillow
pip install pandas
pip install networkx
pip install matplotlib
pip install numpy
pip install scipy
```

### Optional Dependencies

For enhanced functionality:
```bash
pip install openpyxl  # For Excel export
pip install reportlab  # For PDF generation
```

## 🚀 Quick Start

### Option 1: Standalone Executables (Recommended)

1. **No installation required** - just run the executables
2. **Launch the main interface**:
   ```
   _Main.exe
   ```
3. **Select a tool** from the graphical interface and click "Run"
4. **Results** are automatically saved in the `Output` folder

### Option 2: Python Scripts

1. **Clone or download** the repository to your local machine
2. **Install dependencies** using pip (see above)
3. **Run the main interface**:
   ```bash
   python _Main.py
   ```
4. **Select a tool** from the graphical interface and click "Run"

## 📦 Standalone Executables

The following Windows executables are provided for easy deployment:

- **`_Main.exe`** - Main launcher interface (8.2 MB)
- **`0-BID.exe`** - BID Phase assessment tool (36.8 MB)
- **`1-Risk_Assessment_0-A.exe`** - Preliminary risk assessment (15.7 MB)
- **`2-Risk_Assessment.exe`** - Complete risk assessment (15.8 MB)
- **`3-attack_graph_analyzer.exe`** - Attack graph analyzer (794 MB)

### 🎯 Executable Features

- **No Python installation required**
- **Self-contained** - all dependencies included
- **Portable** - can be run from any directory
- **Automatic path management** - finds data files relative to executable location
- **Output management** - creates Output folder automatically

## 📊 Tool Descriptions

### 1. BID Phase Tool (`0-BID.py`)

**Purpose**: Calculate risk value of an ITT (Invitation to Tender) from project category

**Key Features**:
- Cybersecurity requirements assessment
- Project category-based risk scoring
- Interactive scoring matrix (4-point scale)
- Automated risk calculation with weighting
- Word document export capability

**How it works**:
1. Evaluates 11 cybersecurity criteria including:
   - Cybersecurity Requirements
   - Security Architecture Constraints
   - Cryptographic Requirements
   - Authentication & Access Control
   - Supply Chain Security
   - Threat Modeling Guidelines
   - Security Compliance References
   - Security Validation Requirements
   - Incident Response Expectations
   - Data Protection and Privacy
   - Cybersecurity Historical Data

2. Each criterion is scored on a 4-point scale (Low to High)
3. Weighted scoring system calculates final risk value
4. Generates detailed assessment report

**Output**: 
- Risk score with category classification
- Detailed Word document report
- Scoring breakdown and recommendations

---

### 2. Risk Assessment 0-A (`1-Risk_Assessment_0-A.py`)

**Purpose**: Preliminary risk assessment for space missions with threat-asset mapping

**Key Features**:
- 11 predefined threat categories
- 9 asset categories (Ground, Space, Link, User segments)
- 5-criteria scoring system
- Risk matrix calculation
- Interactive threat-asset grid
- Export to Word and CSV formats

**Threat Categories**:
- Data Corruption
- Physical/Logical Attack
- Interception/Eavesdropping
- Jamming
- Denial-of-Service
- Masquerade/Spoofing
- Replay
- Software Threats
- Unauthorized Access/Hijacking
- Tainted hardware components
- Supply Chain

**Assessment Criteria**:
1. **Vulnerability Level** - Known vulnerabilities and their mitigation status
2. **Detection Probability** - Likelihood that malicious activities will be detected
3. **Defense Capability** - Comprehensive defense including mitigations, access controls, and privilege requirements
4. **Operational Impact** - Effect on mission operations
5. **Recovery Time** - Time required to restore normal operations

**Output**:
- Threat-asset risk matrix
- Detailed risk scores per threat-asset combination
- Risk level visualization
- Comprehensive assessment report

---

### 3. Risk Assessment (`2-Risk_Assessment.py`)

**Purpose**: Complete risk assessment tool with advanced analysis capabilities

**Key Features**:
- Advanced threat modeling
- Comprehensive asset management
- Multi-criteria risk evaluation
- Statistical analysis and reporting
- Data import/export functionality
- Customizable risk matrices
- Graphical risk visualization

**Advanced Capabilities**:
- Custom threat definition
- Asset relationship mapping
- Risk aggregation algorithms
- Monte Carlo simulation support
- Sensitivity analysis
- Risk trend analysis
- Mitigation strategy planning

**Assessment Process**:
1. Asset identification and categorization
2. Threat analysis and probability assessment
3. Vulnerability evaluation
4. Impact assessment
5. Risk calculation using multiple methodologies
6. Mitigation recommendation
7. Reporting and documentation

**Output**:
- Comprehensive risk assessment report
- Risk dashboard with visualizations
- Mitigation recommendations
- Compliance mapping
- Executive summary

#### **🛡️ Controls Management** (Enhanced Feature)

**Purpose**: Advanced security controls management with dynamic impact analysis and intelligent asset compatibility

**Key Features**:
- **Dynamic Control Selection**: Searchable control library organized by clusters
- **Smart Asset Filtering**: Intelligent control application based on asset segments (Ground, Space, Link, User, Human Resources)
- **Real-time Impact Analysis**: Live visualization of control effectiveness on threat criteria
- **Threat Coverage Mapping**: Comprehensive analysis of which threats are addressed by selected controls
- **Criteria Impact Visualization**: Shows how controls improve specific threat assessment criteria

**Enhanced Capabilities** (Added 5 August 2025):
- **Segment-aware Control Application**: Controls automatically filter based on asset compatibility
- **Mouse Wheel Protection**: Prevents accidental changes to checkboxes and comboboxes during scrolling
- **Dynamic Search Functionality**: Real-time filtering of controls by title, description, threats, or criteria
- **Expandable Control Clusters**: Organized control categories with collapsible sections
- **Impact Dashboard**: Right-panel analysis showing control effectiveness and threat coverage

**Control Assessment Process**:
1. **Browse Available Controls**: Search and filter from comprehensive control library
2. **Select Relevant Controls**: Choose controls applicable to your assets and threats
3. **Review Dynamic Impact**: See real-time analysis of control effectiveness
4. **Apply to Risk Assessment**: Automatically update threat scores based on selected controls
5. **Export Enhanced Report**: Generate reports including control implementation details

**Smart Features**:
- **Asset Compatibility Checking**: Controls are automatically matched to compatible asset segments
- **Threat Criteria Mapping**: Shows which specific threat criteria are improved by each control
- **Coverage Analysis**: Identifies threats with excellent, good, or basic control coverage
- **Cluster Organization**: Controls grouped by logical categories (e.g., "Access Control", "Encryption", "Monitoring")

**Output Enhancements**:
- Control impact integrated into risk assessment reports
- Detailed control implementation recommendations
- Asset-specific control filtering and suggestions
- Threat coverage gap analysis

---

### 4. Attack Graph Analyzer (`3-attack_graph_analyzer.py`)

**Purpose**: Advanced analysis of threat relationships in space systems with network graph visualization and attack path discovery

**🔥 Key Features**:
- **Interactive Threat Selection** - GUI-based CSV file selection with validation
- **Advanced Network Analysis** - Graph theory algorithms for threat relationships
- **Multiple Analysis Types** - Centrality, critical paths, attack surface analysis
- **Professional Visualizations** - High-quality PNG graphs with customizable layouts
- **Configurable Path Analysis** - User-defined source-target threat combinations
- **Statistical Reporting** - Comprehensive analysis reports with metrics

**🎯 Analysis Capabilities**:

#### **Network Analysis**
- **Graph Statistics**: Nodes, edges, density, connectivity metrics
- **Category Analysis**: Threat distribution by category with visualizations
- **Centrality Analysis**: Degree, betweenness, closeness, PageRank centrality
- **Attack Surface Analysis**: Entry points and final targets identification

#### **Path Analysis**
- **Specific Path Finding**: Discover attack paths between selected threats
- **Multiple Path Analysis**: Batch analysis of predefined threat combinations
- **Critical Path Identification**: High-risk attack sequences with scoring
- **Path Visualization**: Combined graphs showing all discovered paths

#### **Threat Connection Analysis**
- **Star Graph Analysis**: Complete connection map for specific threats
- **Predecessor/Successor Analysis**: Threats that enable or are enabled by target
- **Second-level Neighbors**: Extended threat network exploration
- **Centrality Scoring**: Quantitative importance metrics for individual threats

**🎨 Visualization Features**:
- **Full Network Graph**: Complete threat relationship visualization
- **Star Connection Maps**: Focused analysis around specific threats
- **Attack Path Graphs**: Step-by-step attack sequence visualization
- **Combined Path Views**: Multiple attack paths in single visualization
- **Professional Layouts**: Hierarchical, spring, and custom layouts

**⚙️ Configuration Options**:
```python
# Main path analysis
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": "Social Engineering",
    "target_threat": "Seizure of control: Satellite bus", 
    "max_path_length": 5
}

# Multiple paths analysis
MULTIPLE_PATH_ANALYSIS = [
    {
        "source": "Unauthorized physical access",
        "target": "Denial of Service: Satellite bus", 
        "description": "Physical to DoS attack path"
    }
]

# Analysis parameters
ANALYSIS_PARAMETERS = {
    "top_centrality_nodes": 10,
    "top_critical_paths": 15,
    "max_paths_per_analysis": 20
}
```

**📊 Input Requirements**:
CSV file with threat data containing:
- **THREAT**: Threat name/description
- **Likelihood**: Threat likelihood (Very Low, Low, Medium, High, Very High)
- **Impact**: Threat impact (Very Low, Low, Medium, High, Very High)  
- **Risk**: Overall risk level (Very Low, Low, Medium, High, Very High)

**📈 Output Formats**:
- **Analysis Reports**: Detailed text reports with statistics and findings
- **Network Visualizations**: High-resolution PNG images with legends
- **GEXF Export**: Gephi-compatible graph files for advanced analysis
- **Statistical Summaries**: Quantitative metrics and rankings

**🔍 Advanced Features**:
- **Dynamic Threat Selection**: Runtime CSV file selection with validation
- **Path Criticality Scoring**: Weighted algorithms considering multiple factors
- **Risk-based Highlighting**: Visual emphasis on high-risk threats and paths
- **Relationship Type Analysis**: Different edge types (Enables, Causes, Leads-to)
- **Category-based Filtering**: Focus analysis on specific threat categories

---

## 🔬 Attack Graph Analyzer - Advanced Features

### Interactive Analysis Workflow

1. **Launch Tool**: Run `3-attack_graph_analyzer.exe` or `python 3-attack_graph_analyzer.py`
2. **Select Threat Data**: GUI dialog to choose CSV file with threat information
3. **Automatic Validation**: Tool validates CSV format and required columns
4. **Complete Analysis**: Runs all analysis types automatically
5. **Review Results**: Generated reports and visualizations in Output folder

### Analysis Types Performed

#### **1. Graph Statistics**
```
🚀 GRAPH STATISTICS
Total threats (nodes): 156
Total relationships (edges): 312  
Graph density: 0.026
Average degree: 4.0
```

#### **2. Category Analysis**
- Distribution of threats by category (NAA, EIH, PA, etc.)
- Category relationship matrix
- Visual breakdown with charts

#### **3. Centrality Analysis**
- **Degree Centrality**: Most connected threats
- **Betweenness Centrality**: Bridge threats between different clusters  
- **Closeness Centrality**: Threats with shortest paths to others
- **PageRank**: Most "important" threats in the network

#### **4. Critical Path Analysis**
```
🚨 TOP 15 CRITICAL PATHS IDENTIFIED:
🔥 CRITICAL PATH #1 (Score: 12.45, Danger: 0.89, Length: 4)
   From: Social Engineering
   To: Seizure of control: Satellite bus
   Sequence:
     1. [NAA] Social Engineering
        --(Enables)--> [EIH] Unauthorized access
     2. [EIH] Unauthorized access  
        --(Leads-to)--> [SEC] Security services failure
     3. [SEC] Security services failure
        --(Causes)--> [NAA] Seizure of control: Satellite bus
```

#### **5. Attack Surface Analysis**
- **Entry Points**: Threats with few inputs but many outputs (attack vectors)
- **Final Targets**: Threats with many inputs but few outputs (attack goals)
- Risk-weighted importance scoring

#### **6. Threat Connection Analysis**
```
🔍 CONNECTION ANALYSIS FOR: 'Social Engineering'
📊 BASIC INFORMATION:
   Category: NAA
   Incoming connections: 2
   Outgoing connections: 8
   Total connections: 10

🔽 PREDECESSORS (2) - Threats that LEAD TO 'Social Engineering'
🔼 SUCCESSORS (8) - Threats ENABLED BY 'Social Engineering'
```

### Configuration Examples

#### **Custom Path Analysis**
```python
# Single path analysis
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": "Supply Chain",
    "target_threat": "Data modification",
    "max_path_length": 6
}

# Multiple predefined paths
MULTIPLE_PATH_ANALYSIS = [
    {
        "source": "Physical access",
        "target": "Firmware corruption",
        "description": "Physical to firmware attack"
    },
    {
        "source": "Social Engineering", 
        "target": "Data exfiltration",
        "description": "Social to data theft"
    }
]
```

#### **Visualization Settings**
```python
# Star graph configuration
STAR_GRAPH_CONFIG = {
    "center_threat": "Unauthorized access",
    "max_connections": 20,
    "save_individual": True
}

# Threat connection analysis
THREAT_CONNECTION_ANALYSIS = {
    "save_visualization": True,
    "max_distance": 2,
    "include_predecessors": True,
    "include_successors": True,
    "show_relation_types": True
}
```

### Output Files Generated

#### **Reports**
- `attack_graph_analysis_[timestamp].txt` - Complete analysis report
- Statistical summaries and findings
- Threat rankings and scores

#### **Visualizations** 
- `attack_graph.png` - Full network graph
- `threat_connections_[threat]_[timestamp].png` - Individual threat analysis
- `paths_combined_[source]_[target].png` - Attack path visualization
- Custom graphs based on configuration

#### **Data Export**
- `attack_graph.gexf` - Gephi-compatible graph file
- CSV exports of analysis results
- JSON configuration backups

---

## 📁 File Structure

```
CRAALSP/
├── 📱 EXECUTABLES
│   ├── _Main.exe                     # Main launcher interface (8.2 MB)
│   ├── 0-BID.exe                     # BID Phase assessment (36.8 MB)
│   ├── 1-Risk_Assessment_0-A.exe     # Preliminary risk assessment (15.7 MB)
│   ├── 2-Risk_Assessment.exe         # Complete risk assessment (15.8 MB)
│   └── 3-attack_graph_analyzer.exe   # Attack graph analyzer (794 MB)
│
├── 🐍 PYTHON SCRIPTS
│   ├── _Main.py                      # Main launcher interface
│   ├── 0-BID.py                      # BID Phase assessment tool
│   ├── 1-Risk_Assessment_0-A.py      # Preliminary risk assessment
│   ├── 2-Risk_Assessment.py          # Complete risk assessment
│   ├── 3-attack_graph_analyzer.py    # Attack graph analyzer
│   └── export_import_functions.py    # Shared export/import utilities
│
├── 📋 DATA FILES
│   ├── Asset.json                    # Asset definitions and categories
│   ├── Control.csv                   # Control measures database
│   ├── Legacy.csv                    # Legacy system threat data
│   ├── Threat.csv                    # Comprehensive threat definitions
│   └── attack_graph_threat_relations.csv  # Threat relationships matrix
│
├── 📄 DOCUMENTATION
│   ├── README.md                     # This comprehensive guide
│   └── requirements.txt              # Python dependencies
│
└── 📂 OUTPUT (Auto-generated)
    ├── *.docx                        # Generated assessment reports
    ├── *.png                         # Network visualizations
    ├── *.txt                         # Analysis reports
    ├── *.gexf                        # Gephi graph files
    └── *.csv                         # Exported data
```

### 🔧 Key Data Files

#### **attack_graph_threat_relations.csv**
Defines relationships between threats with columns:
- Source Threat, Target Threat, Relation Type, Source Category, Target Category

#### **Threat.csv** 
Complete threat database with:
- THREAT, Likelihood, Impact, Risk, Category, Description

#### **Asset.json**
Asset definitions for risk assessment including:
- Ground segment, Space segment, Link segment, User segment assets

#### **Control.csv**
Security control measures with effectiveness ratings

## 💡 Usage Examples

### Using the Main Interface

#### **Executable Version (Recommended)**
```bash
# Launch main interface
_Main.exe

# Or run individual tools directly
0-BID.exe
1-Risk_Assessment_0-A.exe  
2-Risk_Assessment.exe
3-attack_graph_analyzer.exe
```

#### **Python Script Version**
```bash
# Launch main interface  
python _Main.py

# Or run individual tools
python 0-BID.py
python 1-Risk_Assessment_0-A.py
python 2-Risk_Assessment.py
python 3-attack_graph_analyzer.py
```

### Main Interface Features

The `_Main.exe` launcher provides:
- **Tool Selection**: Click buttons to launch any of the 4 tools
- **Status Monitoring**: Real-time feedback on tool execution
- **Error Handling**: Graceful error reporting and recovery
- **Unified Interface**: Consistent experience across all tools

### Example Analysis Workflow

#### **1. Space Mission Risk Assessment**
```bash
# Step 1: Initial project assessment
0-BID.exe
# → Evaluates cybersecurity requirements by project category
# → Generates risk score and Word report

# Step 2: Preliminary threat analysis  
1-Risk_Assessment_0-A.exe
# → Maps threats to assets across mission segments
# → Creates threat-asset risk matrix

# Step 3: Comprehensive assessment
2-Risk_Assessment.exe  
# → Detailed risk evaluation with controls
# → Advanced statistical analysis

# Step 4: Attack graph analysis
3-attack_graph_analyzer.exe
# → Network analysis of threat relationships
# → Attack path discovery and visualization
```

#### **2. Attack Graph Analysis Example**

1. **Launch** `3-attack_graph_analyzer.exe`
2. **Select threat data** - GUI opens to choose CSV file
3. **Review validation** - Tool confirms file format
4. **Wait for analysis** - Complete analysis runs automatically
5. **Check results** - Reports and graphs saved to Output folder

**Sample Output Structure:**
```
Output/
├── attack_graph_analysis_20250715_143022.txt
├── attack_graph.png  
├── threat_connections_Social_Engineering_20250715_143025.png
├── paths_combined_Social_Engineering_Seizure_of_control.png
└── attack_graph.gexf
```

### Advanced Attack Graph Usage

#### **Custom Path Analysis**
To analyze specific threat combinations, modify the configuration in the script:

```python
# Edit these variables in 3-attack_graph_analyzer.py
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": "Supply Chain",        # Starting threat
    "target_threat": "Data corruption",     # Target threat  
    "max_path_length": 5                    # Maximum hops
}

MULTIPLE_PATH_ANALYSIS = [
    {
        "source": "Physical access",
        "target": "Denial of Service",
        "description": "Physical to DoS analysis"
    },
    {
        "source": "Social Engineering",
        "target": "Data exfiltration", 
        "description": "Social to data theft"
    }
]
```

#### **Visualization Customization**
```python
# Star graph around specific threat
STAR_GRAPH_CONFIG = {
    "center_threat": "Unauthorized access",
    "max_connections": 15,
    "save_individual": True
}

# Connection analysis settings  
THREAT_CONNECTION_ANALYSIS = {
    "save_visualization": True,
    "max_distance": 3,
    "show_relation_types": True
}
```

### Report Generation Examples

#### **BID Phase Report Output**
```
📊 CYBERSECURITY RISK ASSESSMENT - BID PHASE
Project Category: Communication Satellite (GEO)
Assessment Date: 2025-07-15

🎯 RISK SCORE: 8.2/10 (HIGH RISK)

📋 CRITERIA EVALUATION:
✓ Cybersecurity Requirements: High (4/4)
✓ Security Architecture: Medium (3/4) 
✓ Cryptographic Requirements: High (4/4)
...

📈 RECOMMENDATIONS:
• Implement additional access controls
• Enhance monitoring capabilities
• Review supply chain security
```

#### **Attack Graph Analysis Report**
```
🚀 ATTACK GRAPH ANALYSIS REPORT
Analysis Date: 2025-07-15 14:30:22
Threat Data: Threat_Space_Systems.csv

📊 NETWORK STATISTICS:
• Total Threats: 156 nodes
• Relationships: 312 edges  
• Network Density: 0.026
• Average Connections: 4.0

🎯 TOP CRITICAL PATHS:
1. Social Engineering → Unauthorized access → Security failure → Control seizure
   Score: 12.45, Danger: 0.89, Length: 4 steps

🔍 KEY THREATS BY CENTRALITY:
• Highest Degree: "Unauthorized access" (12 connections)
• Highest Betweenness: "Security services failure" (0.156)
• Highest PageRank: "Social Engineering" (0.034)
```

## 🔧 Configuration

### Attack Graph Analyzer Configuration

The attack graph analyzer supports extensive customization through configuration variables:

#### **Path Analysis Settings**
```python
# Single main path analysis
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": "Social Engineering",           # Starting threat
    "target_threat": "Seizure of control: Satellite bus",  # Target threat
    "max_path_length": 5                            # Maximum path length
}

# Multiple predefined path analyses  
MULTIPLE_PATH_ANALYSIS = [
    {
        "source": "Unauthorized physical access",
        "target": "Denial of Service: Satellite bus",
        "description": "Physical to DoS attack sequence"
    },
    {
        "source": "Supply Chain",  
        "target": "Firmware corruption",
        "description": "Supply chain to firmware attack"
    }
]
```

#### **Analysis Parameters**
```python
ANALYSIS_PARAMETERS = {
    "top_centrality_nodes": 10,      # Number of top central nodes to show
    "top_critical_paths": 15,        # Number of critical paths to analyze
    "max_paths_per_analysis": 20,    # Limit paths per source-target pair
    "path_criticality_threshold": 5.0 # Minimum score for critical paths
}
```

#### **Visualization Settings**
```python
# Star graph configuration (threat-centered analysis)
STAR_GRAPH_CONFIG = {
    "center_threat": "Social Engineering",  # Central threat to analyze
    "max_connections": 20,                  # Maximum connections to show
    "save_individual": True,                # Save individual path graphs
    "layout_type": "spring"                 # Graph layout algorithm
}

# Threat connection analysis
THREAT_CONNECTION_ANALYSIS = {
    "save_visualization": True,       # Save connection visualizations
    "max_distance": 2,               # Include threats up to N hops away
    "include_predecessors": True,     # Include threat predecessors  
    "include_successors": True,       # Include threat successors
    "show_relation_types": True       # Show edge labels with relation types
}
```

#### **File and Output Settings**
```python
# Default threat file (can be overridden by GUI selection)
THREAT_FILE_NAME = "Threat.csv"

# Output configuration
OUTPUT_SETTINGS = {
    "save_full_report": True,        # Save complete analysis report
    "save_individual_graphs": True,  # Save separate graph files
    "image_format": "png",           # Image format (png, svg, pdf)
    "image_dpi": 300,               # Image resolution
    "figure_size": (20, 15)         # Graph dimensions in inches
}
```

### Risk Assessment Configuration

#### **BID Phase Scoring Weights**
```python
# Modify scoring weights in 0-BID.py
CRITERIA_WEIGHTS = {
    "cybersecurity_requirements": 1.2,
    "security_architecture": 1.0, 
    "cryptographic_requirements": 1.1,
    "authentication_access": 1.0,
    "supply_chain_security": 1.3,
    # ... additional criteria
}
```

#### **Risk Matrix Customization**
Risk levels and thresholds can be customized in individual tools by modifying the risk calculation functions.

### Data File Configuration

#### **Threat Relationship Format**
The `attack_graph_threat_relations.csv` file should follow this structure:
```csv
Source Threat,Target Threat,Relation Type,Source Category,Target Category
Social Engineering,Unauthorized access,Enables,NAA,EIH
Unauthorized access,Security services failure,Leads-to,EIH,SEC
Security services failure,Seizure of control,Causes,SEC,NAA
```

#### **Threat Data Format**  
Individual threat CSV files should include:
```csv
THREAT;Likelihood;Impact;Risk
Social Engineering;High;Very High;Very High
Unauthorized physical access;Medium;High;High
Supply Chain;Low;Very High;High
```

### Advanced Configuration

#### **Network Analysis Algorithms**
```python
# Centrality calculation parameters
CENTRALITY_CONFIG = {
    "degree_centrality": True,
    "betweenness_centrality": True, 
    "closeness_centrality": True,
    "pagerank_centrality": True,
    "eigenvector_centrality": False  # Disable for large graphs
}

# Path finding algorithms
PATH_ANALYSIS_CONFIG = {
    "algorithm": "all_simple_paths",    # or "shortest_path"
    "cutoff_length": 6,                 # Maximum path length
    "max_paths_per_pair": 10,          # Limit for performance
    "include_cycles": False             # Exclude circular paths
}
```

## 📊 Data Formats

### Threat Data Format (CSV)
```csv
THREAT;Likelihood;Impact;Risk
Social Engineering;High;Very High;Very High
Unauthorized physical access;Medium;Very High;High
Denial of Service;Low;Medium;Low
```

### Asset Data Format (JSON)
```json
{
  "assets": [
    {
      "name": "Ground Station",
      "category": "Ground",
      "criticality": "High",
      "description": "Primary ground communication facility"
    }
  ]
}
```

## 🎯 Key Features

### User Interface
- **Modern GUI**: Clean, intuitive interfaces using tkinter
- **Responsive Design**: Adapts to different screen resolutions
- **Interactive Elements**: Dynamic forms, real-time calculations
- **Status Monitoring**: Progress tracking and error handling
- **🆕 Integrated Help System**: Comprehensive help dialogs with detailed user guides

### Enhanced User Experience (August 2025)
- **Smart Mouse Wheel Handling**: Prevents accidental value changes during scrolling
- **Dynamic Search and Filtering**: Real-time content filtering with immediate results
- **Interactive Help Dialogs**: Built-in comprehensive user guides accessible via ❓ Help buttons
- **Asset Compatibility Intelligence**: Automatic filtering of relevant controls based on asset segments
- **Real-time Impact Visualization**: Live updates of control effectiveness and threat coverage

### Analysis Capabilities
- **Multi-criteria Assessment**: Comprehensive evaluation frameworks
- **Statistical Analysis**: Advanced mathematical models
- **Visualization**: Charts, graphs, and network diagrams
- **Reporting**: Professional document generation

### Export/Import Functions
- **Word Documents**: Formatted reports with tables and charts
- **CSV Files**: Raw data for further analysis
- **JSON Format**: Structured data exchange
- **PNG Images**: High-quality visualizations

### 🆕 Help and Documentation System
- **Interactive Help Dialogs**: Built-in comprehensive user guides
- **Context-sensitive Help**: Help buttons in complex interfaces (Controls Management)
- **Feature Explanations**: Detailed descriptions of enhanced features and capabilities
- **Best Practices Guidance**: Tips for optimal tool usage and configuration
- **Updated Documentation**: README includes all latest features and enhancements

## 🔍 Troubleshooting

### Common Issues

1. **Import Errors**: Install missing dependencies using pip
2. **File Not Found**: Ensure CSV files are in the correct directory
3. **Permission Errors**: Run with appropriate file system permissions
4. **Memory Issues**: Large datasets may require increased memory allocation

### Error Messages

- **"File not found"**: Check file paths and ensure CSV files exist
- **"Invalid data format"**: Verify CSV file structure matches requirements
- **"Export failed"**: Check write permissions in output directory

## 📈 Performance Notes

- **Large Datasets**: Tools are optimized for typical space mission scenarios
- **Memory Usage**: Network analysis may require significant RAM for large graphs
- **Processing Time**: Complex analyses may take several minutes to complete

## 🔒 Security Considerations

- **Data Privacy**: Assessment data should be handled according to organizational policies
- **File Security**: Ensure proper access controls on assessment files
- **Network Security**: Consider security implications of threat data sharing

## 🤝 Contributing

This tool suite was developed as part of academic research for space program cybersecurity assessment. The project is open for contributions and enhancements.

### Development Guidelines

#### **Code Standards**
- Follow PEP 8 Python style guidelines
- Include comprehensive docstrings for all functions
- Maintain backward compatibility where possible
- Add unit tests for new functionality

#### **Contributing Process**
1. **Fork** the repository from GitHub
2. **Create** a feature branch (`git checkout -b feature/new-analysis`)
3. **Implement** your changes with appropriate documentation
4. **Test** thoroughly with sample data
5. **Submit** a pull request with detailed description

#### **Areas for Contribution**
- **New Analysis Algorithms**: Additional graph analysis methods
- **Visualization Improvements**: Enhanced plotting and export options
- **Data Format Support**: Additional input/output file formats
- **Performance Optimization**: Scalability for larger datasets
- **User Interface**: GUI enhancements and usability improvements

#### **Research Extensions**
- Integration with other cybersecurity frameworks
- Machine learning-based threat prediction
- Real-time threat monitoring capabilities
- Integration with space system simulators

### Academic Usage

#### **Citation**
When using this tool suite in academic work, please cite:
```
Nonni, G. (2025). Cyber Risk Assessment Across Lifecycle of Space Program (CRAALSP). 
Thesis work, [University Name], Student ID: 1948023.
```

#### **Research Applications**
- Space mission security assessment
- Cybersecurity risk analysis methodologies
- Attack graph analysis in space systems
- Threat modeling for satellite operations

## 📜 License

### Academic License

This project is developed as part of thesis work for space program risk assessment. Usage terms:

- **Academic Use**: Free for educational and research purposes
- **Commercial Use**: Contact author for licensing arrangements
- **Modification**: Permitted with attribution to original work
- **Distribution**: Allowed with proper academic citation

### Disclaimer

This tool suite is provided for educational and research purposes. While developed with space industry best practices:

- **No Warranty**: Tools provided "as-is" without guarantees
- **Validation Required**: Results should be validated by domain experts
- **Responsibility**: Users responsible for appropriate application
- **Security**: Ensure proper handling of sensitive assessment data

### Third-Party Licenses

This project incorporates open-source libraries with their respective licenses:
- **Python**: PSF License
- **NetworkX**: BSD License  
- **Matplotlib**: PSF-based License
- **Pandas**: BSD License
- **NumPy**: BSD License

## 👤 Author & Contact

### Project Author
**Giuseppe Nonni**  
- **Student ID**: 1948023
- **Email**: giuseppe.nonni@gmail.com
- **Institution**: [University Name]
- **Field**: Space Systems Engineering, Cybersecurity

### Thesis Information
- **Title**: Cyber Risk Assessment Across Lifecycle of Space Program
- **Focus**: Cybersecurity methodologies for space missions
- **Year**: 2025
- **Supervisor**: [Supervisor Name]

### Technical Support

#### **For Tool Issues**
- **GitHub Issues**: Report bugs and feature requests
- **Email Support**: Technical questions and implementation help
- **Documentation**: Check README and source code comments

#### **For Research Questions**
- **Space Domain Expertise**: Contact for space-specific cybersecurity questions
- **Methodology Discussion**: Available for academic collaboration
- **Tool Validation**: Assistance with tool validation and verification

#### **Acknowledgments**
Special thanks to:
- Thesis supervisor for guidance and support
- Space industry experts for domain knowledge validation
- Open-source community for foundational libraries
- Academic reviewers for feedback and improvements

---

## 📚 Additional Resources

### Related Documentation
- **Space Cybersecurity Standards**: ISO/IEC 27001, NIST Cybersecurity Framework
- **Attack Graph Theory**: Academic papers on graph-based security analysis
- **Space Mission Security**: ESA and NASA cybersecurity guidelines

### Further Reading
- Space Systems Security Engineering methodologies
- Graph theory applications in cybersecurity
- Risk assessment frameworks for critical infrastructure
- Satellite constellation security considerations

### Tool Integration
- **Gephi**: For advanced graph visualization and analysis
- **MITRE ATT&CK**: Framework integration for threat mapping
- **Risk Management Tools**: Integration with enterprise risk platforms

---

*This tool suite represents a comprehensive approach to cybersecurity risk assessment in space missions. It combines academic research with practical implementation to support the growing need for space system security analysis.*

**🚀 CRAALSP - Securing the Future of Space Exploration 🛰️**
