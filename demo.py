#!/usr/bin/env python3
"""
🚀 Demo Script for AI-Powered TOD Control Analysis Tool

This interactive demonstration showcases the capabilities of the AI-driven
internal control analysis framework. Perfect for understanding the tool's
functionality before analyzing your own control data.

Author: Mehul Chhabra
GitHub: https://github.com/Mehulchhabra07/Automated-Test-of-Design---Internal-Controls
"""

import os
import sys
from pathlib import Path

def print_banner():
    """Display an attractive banner for the demo"""
    banner = """
    ╔══════════════════════════════════════════════════════════════╗
    ║                🤖 AI-Powered TOD Control Analysis             ║
    ║                        Demo Application                      ║
    ║                                                              ║
    ║   Transform your internal control testing with AI! 🚀       ║
    ╚══════════════════════════════════════════════════════════════╝
    """
    print(banner)

def setup_demo():
    """Set up the demo environment with comprehensive checks"""
    print_banner()
    print("� Performing environment validation...\n")
    
    # Check Python version
    python_version = sys.version_info
    if python_version < (3, 8):
        print("❌ Python 3.8+ required. Current version:", 
              f"{python_version.major}.{python_version.minor}")
        return False
    print(f"✅ Python {python_version.major}.{python_version.minor} detected")
    
    # Check dependencies
    required_packages = ['pandas', 'openpyxl', 'openai', 'httpx']
    missing_packages = []
    
    for package in required_packages:
        try:
            __import__(package)
            print(f"✅ {package} installed")
        except ImportError:
            missing_packages.append(package)
            print(f"❌ {package} missing")
    
    if missing_packages:
        print(f"\n⚠️  Missing packages: {', '.join(missing_packages)}")
        print("Please install them with: pip install -r requirements.txt")
        return False
    
    # Check OpenAI API key
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        print("\n⚠️  OpenAI API key not found!")
        print("   Set your API key with:")
        print("   export OPENAI_API_KEY='your-api-key-here'")
        print("\n   Or for this session only:")
        api_key = input("   Enter your OpenAI API key (or press Enter to skip): ").strip()
        
        if api_key:
            os.environ["OPENAI_API_KEY"] = api_key
            print("   ✅ API key set for this session")
        else:
            print("   ⚠️  Continuing without API key (analysis will fail)")
            return False
    else:
        masked_key = f"{api_key[:7]}...{api_key[-4:]}" if len(api_key) > 11 else "***"
        print(f"✅ OpenAI API key found: {masked_key}")
    
    # Check if sample file exists
    sample_file = Path("sample_controls.xlsx")
    if not sample_file.exists():
        print("\n📄 Sample file not found, creating one...")
        if create_sample_file():
            print("   ✅ Sample file created successfully")
        else:
            print("   ❌ Failed to create sample file")
            return False
    else:
        print("✅ Sample file 'sample_controls.xlsx' found")
    
    print("\n🎉 Demo environment ready!")
    print("=" * 60)
    return True

def create_sample_file():
    """Create a comprehensive sample Excel file for demonstration"""
    try:
        import pandas as pd
        
        # Enhanced sample data with more realistic examples
        data = {
            'Risk': ['R001', 'R002', 'R003', 'R004'],
            'Risk Description': [
                'Risk of unauthorized access to sensitive financial data resulting in data breaches, fraud, or regulatory violations',
                'Risk of erroneous financial reporting due to manual data entry errors, system glitches, and lack of validation controls', 
                'Risk of incomplete expense approvals leading to unauthorized payments, budget overruns, and fraud',
                'Risk of inadequate data backup and recovery procedures resulting in data loss during system failures or cyber attacks'
            ],
            'Control': ['C001', 'C002', 'C003', 'C004'],
            'Control Description': [
                'The IT Security Manager performs monthly comprehensive review of user access privileges including role verification, dormant account identification, access rights validation, and segregation of duties compliance in the SAP financial system',
                'Automated system validation checks are performed in real-time on all financial entries with exception reporting to the Finance Manager, including data type validation, range checks, duplicate detection, and business rule verification',
                'Department heads review and approve all expenses above $1,000 using digital approval workflow with dual authorization requirement, documented business justification, and budget availability verification',
                'IT team performs weekly automated backups of critical financial data with monthly restore testing, quarterly disaster recovery drills, and annual business continuity plan review'
            ],
            'Automation': ['Manual', 'Automated', 'Semi-Auto', 'Automated'],
            'Detective/ Preventive': ['Detective', 'Preventive', 'Preventive', 'Preventive'],
            'Operation Frequency': ['Monthly', 'Real-time', 'As needed', 'Weekly']
        }
        
        df = pd.DataFrame(data)
        df.to_excel('sample_controls.xlsx', index=False, engine='openpyxl')
        
        print(f"   📊 Created {len(df)} sample controls")
        print(f"   📋 Columns: {', '.join(df.columns)}")
        return True
        
    except ImportError:
        print("   ❌ pandas not installed. Please run: pip install pandas openpyxl")
        return False
    except Exception as e:
        print(f"   ❌ Error creating sample file: {e}")
        return False

def display_analysis_preview():
    """Show what the analysis will evaluate"""
    print("\n🔍 ANALYSIS PREVIEW")
    print("=" * 60)
    print("The AI will evaluate each control across 9 dimensions:")
    print()
    
    dimensions = [
        ("🎯 Completeness", "6W Framework (Who, What, When, Where, Why, How)"),
        ("🛡️  Control Objective", "Does the control mitigate the identified risk?"),
        ("⚙️  Execution", "Is the automation level appropriate?"),
        ("🏷️  Type Adequacy", "Detective vs Preventive classification"),
        ("⏰ Frequency", "Is the operating frequency suitable?"),
        ("🖥️  Dependencies", "What systems and tools are involved?"),
        ("👥 Segregation", "Are duties properly separated?"),
        ("📊 Overall Rating", "Effective, Partially Effective, or Ineffective"),
        ("📋 Evidence", "What should auditors look for?")
    ]
    
    for dimension, description in dimensions:
        print(f"   {dimension}: {description}")
    
    print("\n📈 EXPECTED OUTPUT")
    print("=" * 60)
    print("📄 Excel report with:")
    print("   • Color-coded completeness analysis")
    print("   • Detailed AI explanations for each assessment")
    print("   • Improvement suggestions for gaps")
    print("   • Professional formatting and styling")
    print("   • Expected evidence recommendations")

def run_demo():
    """Run the complete demonstration"""
    if not setup_demo():
        print("\n❌ Demo setup failed. Please resolve the issues above.")
        return False
    
    display_analysis_preview()
    
    print("\n🚀 STARTING ANALYSIS")
    print("=" * 60)
    
    # Get user confirmation
    while True:
        response = input("\nProceed with AI analysis? (y/n): ").lower().strip()
        if response in ['y', 'yes']:
            break
        elif response in ['n', 'no']:
            print("Demo cancelled by user.")
            return True
        else:
            print("Please enter 'y' for yes or 'n' for no.")
    
    print("\n🔄 Running control analysis...")
    print("   This may take 2-5 minutes depending on API response times...")
    print("   Please be patient while the AI analyzes each control...")
    print()
    
    try:
        # Import and run the main analysis
        from tod_control_testing_v21_enhanced import main
        main()
        
        print("\n" + "=" * 60)
        print("🎉 DEMO COMPLETED SUCCESSFULLY!")
        print("=" * 60)
        print("📄 Results saved to: sample_controls_TestResult.xlsx")
        print("📊 Open the file to see the detailed AI analysis")
        print()
        print("� Next Steps:")
        print("   1. Review the analysis results in Excel")
        print("   2. Examine the AI's reasoning and suggestions")
        print("   3. Try with your own control data")
        print("   4. Customize the analysis parameters if needed")
        print()
        print("🔗 Learn more: https://github.com/Mehulchhabra07/Automated-Test-of-Design---Internal-Controls")
        
        return True
        
    except KeyboardInterrupt:
        print("\n⚠️  Analysis interrupted by user")
        return False
    except Exception as e:
        print(f"\n❌ Demo failed with error: {e}")
        print("\nTroubleshooting tips:")
        print("   • Check your API key is valid and has credits")
        print("   • Ensure stable internet connection")
        print("   • Try again in a few minutes if rate limited")
        print("   • Check the logs in 'tod_analysis.log' for details")
        return False

def main():
    """Main demo function"""
    try:
        success = run_demo()
        sys.exit(0 if success else 1)
    except KeyboardInterrupt:
        print("\n\n👋 Demo interrupted. Goodbye!")
        sys.exit(1)

if __name__ == "__main__":
    main()
