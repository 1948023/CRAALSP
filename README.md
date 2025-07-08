# Risk Assessment Tool Suite for Space Systems

A comprehensive cybersecurity risk assessment toolkit designed specifically for space missions and satellite systems. This suite provides tools for threat analysis, risk evaluation, and attack graph visualization to support space program security assessments.

## 🚀 Overview

This tool suite was developed as part of a thesis work for space program risk assessment and consists of four main components:

1. **BID Phase Tool** - Initial risk assessment based on project category
2. **Risk Assessment 0-A** - Preliminary threat and vulnerability analysis
3. **Risk Assessment** - Comprehensive risk evaluation with detailed criteria
4. **Attack Graph Analyzer** - Threat relationship analysis and visualization

## 📋 Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [Tool Descriptions](#tool-descriptions)
- [File Structure](#file-structure)
- [Usage Examples](#usage-examples)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [License](#license)

## 🛠️ Installation

### Prerequisites

- Python 3.7 or higher
- Windows/Linux/macOS

### Required Libraries

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

1. **Clone or download** the repository to your local machine
2. **Install dependencies** using pip (see above)
3. **Run the main interface**:
   ```bash
   python _Main.py
   ```
4. **Select a tool** from the graphical interface and click "Run"

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
2. **Access Control** - Physical and logical access protection measures
3. **Defense Capability** - Countermeasures and detection systems
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

---

### 4. Attack Graph Analyzer (`3-attack_graph_analyzer.py`)

**Purpose**: Analyze relationships between threats in space systems and create threat attack graphs

**Key Features**:
- Interactive file selection for threat data
- Network graph visualization
- Path analysis between threats
- Centrality analysis
- Star graph generation
- Multiple visualization formats
- Statistical analysis of threat relationships

**Analysis Capabilities**:
- **Path Analysis**: Find attack paths between specific threats
- **Centrality Analysis**: Identify key threats in the network
- **Star Graph Analysis**: Show all connections to a specific threat
- **Critical Path Identification**: Discover high-risk attack sequences
- **Network Metrics**: Calculate graph statistics and properties

**Visualization Options**:
- Full threat network graph
- Specific threat connection maps
- Attack path visualizations
- Centrality heat maps
- Risk-weighted network layouts

**Configuration Options**:
- Adjustable path length limits
- Customizable threat selection
- Flexible visualization parameters
- Export format selection

**Input Requirements**:
- CSV file with threat data containing:
  - THREAT: Threat name
  - Likelihood: Threat likelihood (Very Low to Very High)
  - Impact: Threat impact (Very Low to Very High)
  - Risk: Overall risk level (Very Low to Very High)

**Output**:
- Network graph visualizations (PNG format)
- Attack path analysis report
- Centrality analysis results
- Statistical summary
- Interactive graph files

---

## 📁 File Structure

```
Risk Assessment Tool Suite/
├── _Main.py                          # Main launcher interface
├── 0-BID.py                         # BID Phase assessment tool
├── 1-Risk_Assessment_0-A.py         # Preliminary risk assessment
├── 2-Risk_Assessment.py             # Complete risk assessment
├── 3-attack_graph_analyzer.py       # Attack graph analyzer
├── export_import_functions.py       # Shared export/import utilities
├── README.md                        # This file
├── Asset.json                       # Asset definitions
├── Control.csv                      # Control measures database
├── Legacy.csv                       # Legacy system data
├── Threat.csv                       # Threat definitions
├── attack_graph_threat_relations.csv # Threat relationships
└── CSV_Export_[timestamp]/          # Export directory
    ├── Threat_Analyzed_*.csv        # Analyzed threat data
    └── [various analysis files]
```

## 💡 Usage Examples

### Running the Main Interface

```bash
python _Main.py
```

This launches the graphical interface where you can:
- Select any of the four tools
- Monitor execution status
- View completion notifications

### Running Individual Tools

Each tool can also be run independently:

```bash
# BID Phase Assessment
python 0-BID.py

# Preliminary Risk Assessment
python 1-Risk_Assessment_0-A.py

# Complete Risk Assessment
python 2-Risk_Assessment.py

# Attack Graph Analyzer
python 3-attack_graph_analyzer.py
```

### Example Workflow

1. **Start with BID Phase** to assess initial project risk
2. **Use Risk Assessment 0-A** for preliminary threat analysis
3. **Run complete Risk Assessment** for detailed evaluation
4. **Analyze threat relationships** with Attack Graph Analyzer
5. **Generate comprehensive reports** from all tools

## 🔧 Configuration

### Attack Graph Analyzer Configuration

The attack graph analyzer can be configured by modifying variables in the script:

```python
# File selection (interactive or programmatic)
THREAT_FILE_NAME = "CSV_Export_[timestamp]/Threat_Analyzed.csv"

# Path analysis configuration
SPECIFIC_PATH_ANALYSIS = {
    "source_threat": "Social Engineering",
    "target_threat": "Seizure of control: Satellite bus",
    "max_path_length": 5
}

# Visualization settings
save_path = 0  # Save individual path plots
max_five = 0   # Limit to 5 combined paths
```

### Risk Assessment Configuration

Risk matrices and scoring criteria can be customized in each tool by modifying the relevant data structures in the source code.

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

This tool suite was developed as part of academic research. For contributions or modifications:

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Submit a pull request with detailed description

## 📜 License

This project is part of a thesis work for space program risk assessment. Please refer to your institution's policies regarding academic software usage and distribution.

## 👤 Author

**Giuseppe Nonni** (Student ID: 1948023)  
Email: giuseppe.nonni@gmail.com  
Thesis work for space program risk assessment tool

## 📞 Support

For technical support or questions:
- Check the troubleshooting section above
- Review the source code comments for detailed implementation notes
- Contact the author for specific issues

---

*This tool suite is designed to support cybersecurity risk assessment in space missions and satellite systems. It provides a comprehensive framework for threat analysis, risk evaluation, and security planning in the space domain.*
