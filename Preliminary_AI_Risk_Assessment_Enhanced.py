#!/usr/bin/env python3
"""
Preliminary AI Risk Assessment Tool
Generates a preliminary risk analysis in CCSDS 350.1-G-3 style with mission-specific prompts
"""

import pandas as pd
import requests
import sys
import os
from datetime import datetime

CCSDS_THREATS = [
    "Data Corruption",
    "Physical Attack",
    "Interception/Eavesdropping",
    "Jamming",
    "Denial-of-Service",
    "Masquerade/Spoofing",
    "Replay",
    "Software Threats",
    "Unauthorized Access/Hijacking",
    "Tainted Hardware Components",
    "Supply Chain"
]

# Mission types and their characteristics for inference
MISSION_TYPES = {
    "Earth Observation": {
        "keywords": ["earth observation", "remote sensing", "imaging", "monitoring", "surveillance", "environmental", "optical", "radar"],
        "orbit_keywords": ["LEO", "polar", "sun-synchronous", "low earth"]
    },
    "Communication": {
        "keywords": ["communication", "telecommunications", "relay", "broadcasting", "internet", "voice", "data", "constellation"],
        "orbit_keywords": ["GEO", "MEO", "LEO constellation", "geostationary"]
    },
    "Science Mission": {
        "keywords": ["science", "research", "exploration", "astronomy", "astrophysics", "planetary", "deep space", "lunar", "mars"],
        "orbit_keywords": ["lunar", "mars", "deep space", "heliocentric", "interplanetary"]
    },
    "Navigation": {
        "keywords": ["navigation", "positioning", "GPS", "GNSS", "timing", "location", "atomic clock"],
        "orbit_keywords": ["MEO", "medium earth orbit", "navigation"]
    },
    "On-Orbit Service": {
        "keywords": ["servicing", "refueling", "repair", "debris removal", "satellite maintenance", "robotics", "docking"],
        "orbit_keywords": ["various", "multiple orbits", "rendezvous"]
    }
}

# Mission-specific context for prompts
MISSION_CONTEXT = {
    "Earth Observation": {
        "key_assets": "imaging sensors, data processing systems, ground stations, data storage",
        "critical_functions": "Earth imaging, data collection, environmental monitoring",
        "typical_threats": "data theft, image manipulation, unauthorized surveillance"
    },
    "Communication": {
        "key_assets": "transponders, antennas, user terminals, ground gateways",
        "critical_functions": "voice/data relay, internet connectivity, broadcasting",
        "typical_threats": "eavesdropping, jamming, service disruption"
    },
    "Science Mission": {
        "key_assets": "scientific instruments, data recorders, navigation systems",
        "critical_functions": "scientific data collection, instrument control, mission operations",
        "typical_threats": "data corruption, instrument sabotage, mission interference"
    },
    "Navigation": {
        "key_assets": "atomic clocks, signal generators, control systems, user receivers",
        "critical_functions": "precise timing, positioning signals, navigation services",
        "typical_threats": "signal spoofing, timing attacks, navigation disruption"
    },
    "On-Orbit Service": {
        "key_assets": "robotic arms, docking systems, proximity sensors, control systems",
        "critical_functions": "satellite servicing, debris removal, orbital operations",
        "typical_threats": "hijacking, collision, unauthorized maneuvers"
    }
}

ISO_27005_MATRIX = {
    ("very low", "very low"): "very low",
    ("very low", "low"): "very low",
    ("very low", "medium"): "low",
    ("very low", "high"): "medium",
    ("very low", "very high"): "medium",
    ("low", "very low"): "very low",
    ("low", "low"): "low",
    ("low", "medium"): "low",
    ("low", "high"): "medium",
    ("low", "very high"): "medium",
    ("medium", "very low"): "low",
    ("medium", "low"): "low",
    ("medium", "medium"): "medium",
    ("medium", "high"): "high",
    ("medium", "very high"): "high",
    ("high", "very low"): "low",
    ("high", "low"): "medium",
    ("high", "medium"): "high",
    ("high", "high"): "high",
    ("high", "very high"): "very high",
    ("very high", "very low"): "medium",
    ("very high", "low"): "high",
    ("very high", "medium"): "high",
    ("very high", "high"): "very high",
    ("very high", "very high"): "very high"
}

class PreliminaryAIRiskAssessment:
    def __init__(self, description_file):
        self.description_file = description_file
        self.model = "mistral:7b"
        self.ollama_url = "http://localhost:11434"
        self.program_description = self.load_description()
        self.mission_type = self.infer_mission_type()
        self.assets = []

    def load_description(self):
        with open(self.description_file, 'r', encoding='utf-8') as f:
            return f.read().strip()

    def infer_mission_type(self):
        """Infer mission type from program description"""
        description_lower = self.program_description.lower()
        
        scores = {}
        for mission_type, data in MISSION_TYPES.items():
            score = 0
            # Check keywords
            for keyword in data["keywords"]:
                if keyword in description_lower:
                    score += 2
            # Check orbit keywords
            for orbit_keyword in data["orbit_keywords"]:
                if orbit_keyword.lower() in description_lower:
                    score += 3
            scores[mission_type] = score
        
        # Return mission type with highest score, default to Earth Observation
        if max(scores.values()) == 0:
            return "Earth Observation"
        
        inferred_type = max(scores, key=lambda k: scores[k])
        print(f"Inferred mission type: {inferred_type} (score: {scores[inferred_type]})")
        return inferred_type

    def query_ollama(self, prompt):
        try:
            response = requests.post(
                f"{self.ollama_url}/api/generate",
                json={
                    "model": self.model,
                    "prompt": prompt,
                    "stream": False,
                    "options": {
                        "temperature": 0.3,
                        "num_predict": 1500,
                        "num_ctx": 4096,
                        "top_k": 40,
                        "top_p": 0.9
                    }
                },
                timeout=300
            )
            if response.status_code == 200:
                result = response.json().get('response', '').strip()
                return result if result else "Analysis failed - no response from AI model"
            else:
                print(f"HTTP Error: {response.status_code}")
                return f"HTTP Error {response.status_code} - Analysis failed"
        except Exception as e:
            print(f"Query error: {e}")
            return f"Query error: {str(e)} - Analysis failed"

    def analyze_context(self):
        print ("Analysis context")
        mission_context = MISSION_CONTEXT.get(self.mission_type, MISSION_CONTEXT["Earth Observation"])
        
        prompt = f"""You are a cybersecurity analyst for satellite systems. Analyze the following {self.mission_type.lower()} program description and provide:

1. PROGRAM CONTEXT ANALYSIS (mission, environment, criticality)
2. RELEVANT ASSETS IDENTIFICATION (focus on {mission_context['key_assets']})

Mission Type: {self.mission_type}
Key Functions: {mission_context['critical_functions']}
Typical Security Concerns: {mission_context['typical_threats']}

Be concise but complete. Format with clear sections."""
        prompt += f"\n\nPROGRAM DESCRIPTION:\n{self.program_description}"
        return self.query_ollama(prompt)

    def analyze_threats(self):
        print ("Analysis threats")
        mission_context = MISSION_CONTEXT.get(self.mission_type, MISSION_CONTEXT["Earth Observation"])
        results = []
        
        for threat in CCSDS_THREATS:
            prompt = f"""You are a cybersecurity analyst specializing in {self.mission_type.lower()} satellites. 

MISSION TYPE: {self.mission_type}
KEY ASSETS: {mission_context['key_assets']}
CRITICAL FUNCTIONS: {mission_context['critical_functions']}

PROGRAM: {self.program_description}
THREAT: {threat}

For this {threat} threat in the context of {self.mission_type.lower()} missions:
- Describe how it could specifically impact this type of mission
- Consider the typical {mission_context['typical_threats']} for this mission type
- Assign a probability level (very low, low, medium, high, very high) with orbit-specific considerations if relevant
- Assign an impact level (very low, low, medium, high, very high) based on mission criticality
- Calculate risk level using ISO 27005 risk matrix
- Recommend specific security controls appropriate for {self.mission_type.lower()} missions

Format as:
Threat: {threat}
Mission-Specific Impact: [how this threat affects {self.mission_type.lower()} missions]
Probability: [level] ([justification considering mission type and orbit])
Impact: [level] ([justification based on mission criticality])
Risk Level: [level] 
Security Controls: [mission-appropriate controls]"""
            
            analysis = self.query_ollama(prompt)
            results.append(f"## Threat: {threat}\n{analysis}\n")
        return '\n'.join(results)

    def overall_risk_summary(self, threats_analysis):
        print ("Creating summary")
        # Estrai solo i nomi dei threat e i loro risk level per evitare prompt troppo lunghi
        threat_summary = []
        lines = threats_analysis.split('\n')
        current_threat = ""
        current_risk = ""
        
        for line in lines:
            if line.startswith("## Threat:"):
                current_threat = line.replace("## Threat:", "").strip()
            elif "Risk Level:" in line:
                current_risk = line.split("Risk Level:")[1].split("(")[0].strip() if "Risk Level:" in line else "unknown"
                if current_threat and current_risk:
                    threat_summary.append(f"{current_threat}: {current_risk}")
        
        threats_list = "\n".join(threat_summary)
        
        prompt = f"""You are a cybersecurity analyst specializing in {self.mission_type.lower()} satellites. Based on the following satellite program and cybersecurity threat analysis results, provide a concise summary:

MISSION TYPE: {self.mission_type}
PROGRAM: {self.program_description}

CYBERSECURITY THREATS ANALYZED:
{threats_list}

Provide:
- Overall cybersecurity risk level for this {self.mission_type.lower()} program (very low, low, medium, high, very high)
- Top 3 highest cybersecurity risks from the list above
- Key cybersecurity mitigation strategies specific to {self.mission_type.lower()} missions

IMPORTANT: Use ONLY the threats listed above. Focus on {self.mission_type.lower()} mission-specific risks. Be concise."""
        return self.query_ollama(prompt)

    def run_preliminary_assessment(self):
        print("Starting preliminary risk assessment...")
        start = datetime.now().strftime("%Y%m%d_%H%M%S")    
        print(start)
        context = self.analyze_context()
        threats = self.analyze_threats()
        summary = self.overall_risk_summary(threats)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"Output/Preliminary_Risk_Assessment_{self.mission_type.replace(' ', '_')}_{timestamp}.md"
        os.makedirs("Output", exist_ok=True)
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(f"# Preliminary Risk Assessment Report\n\n")
            f.write(f"**Mission Type:** {self.mission_type}\n")
            f.write(f"**Satellite Program:**\n\n{self.program_description}\n\n")
            f.write("---\n\n")
            f.write("## 1. Program Context Analysis & Asset Identification\n\n")
            f.write(context)
            f.write("\n\n---\n\n")
            f.write("## 2. Threat Analysis (CCSDS 350.1-G-3)\n\n")
            f.write(threats)
            f.write("\n\n---\n\n")
            f.write("## 3. Overall Risk Summary\n\n")
            f.write(summary)
            f.write("\n\n---\n")
        print(f"Preliminary report saved to: {filename}")
        return filename

def main():
    print("Preliminary AI Risk Assessment Tool (CCSDS 350.1-G-3)")
    print("=" * 60)
    if len(sys.argv) > 1:
        description_file = sys.argv[1]
    else:
        description_file = input("Enter the path to the program description file: ").strip()
    if not os.path.isfile(description_file):
        print("Description file not found.")
        sys.exit(1)
    assessment = PreliminaryAIRiskAssessment(description_file)
    result = assessment.run_preliminary_assessment()
    if result:
        print(f"\nPreliminary assessment complete! Check the file: {result}")
    else:
        print("\nAssessment failed")

if __name__ == "__main__":
    main()
