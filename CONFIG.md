# 🔧 Configuration Guide

Complete setup and configuration guide for the AI-Powered TOD Control Analysis Tool.

## 🚀 Quick Setup

### 1. Environment Variables
Set up your OpenAI API key (required):

```bash
# For macOS/Linux
export OPENAI_API_KEY="sk-your-openai-api-key-here"

# For Windows
set OPENAI_API_KEY=sk-your-openai-api-key-here

# Or add to your .bashrc/.zshrc for persistence
echo 'export OPENAI_API_KEY="sk-your-openai-api-key-here"' >> ~/.bashrc
```

### 2. Getting Your OpenAI API Key

1. Visit [OpenAI Platform](https://platform.openai.com/api-keys)
2. Sign in or create an account
3. Click "Create new secret key"
4. Copy the key (starts with `sk-`)
5. Set it as an environment variable (see above)

## 📊 Input File Format

Your Excel file must have these **exact column headers**:

| Column Name | Type | Description | Required | Example |
|-------------|------|-------------|----------|---------|
| **Risk** | String | Unique risk identifier | ✅ | R001, RISK-FIN-001 |
| **Risk Description** | String | Detailed risk description | ✅ | Risk of unauthorized access to financial data |
| **Control** | String | Unique control identifier | ✅ | C001, CTRL-ACC-001 |
| **Control Description** | String | Detailed control description | ✅ | Monthly review of user access privileges |
| **Automation** | String | Level of automation | ✅ | Manual, Semi-Auto, Automated |
| **Detective/ Preventive** | String | Control type | ✅ | Detective, Preventive |
| **Operation Frequency** | String | Operating frequency | ✅ | Daily, Weekly, Monthly, Quarterly |

### 📝 Sample Control Description

Here's an example of a **well-formed** control description that includes all 6W elements:

```
The IT Security Manager (WHO) performs a monthly review (WHEN) of all user 
access privileges in the SAP financial system (WHERE) to ensure proper 
segregation of duties and prevent unauthorized access (WHY). The review 
includes verification of user roles, identification of dormant accounts, 
and validation of access rights (WHAT). The process follows a standardized 
checklist with documented approval workflows (HOW).
```

**Elements covered:**
- ✅ **Who**: IT Security Manager
- ✅ **What**: Review of user access privileges  
- ✅ **When**: Monthly
- ✅ **Where**: SAP financial system
- ✅ **Why**: Ensure segregation of duties, prevent unauthorized access
- ✅ **How**: Standardized checklist with approval workflows

## ⚙️ Advanced Configuration

### File Paths
Update these in the `Config` class:

```python
class Config:
    # Input/Output files
    INPUT_FILE = Path("your_controls.xlsx")
    OUTPUT_FILE = Path("analysis_results.xlsx")
    
    # Or use absolute paths
    INPUT_FILE = Path("/path/to/your/controls.xlsx")
    OUTPUT_FILE = Path("/path/to/results/analysis_results.xlsx")
```

### AI Model Selection
Choose the best model for your needs:

```python
# For maximum accuracy (recommended)
MODEL = "gpt-4"

# For faster processing and lower cost
MODEL = "gpt-3.5-turbo"

# For latest features
MODEL = "gpt-4o"
```

**Cost Comparison (approximate per 1K tokens):**
- `gpt-3.5-turbo`: $0.002 (fastest, lowest cost)
- `gpt-4`: $0.03 (best accuracy)
- `gpt-4o`: $0.005 (good balance)

### Analysis Elements
Customize what the AI analyzes:

```python
# Default 6W framework
ELEMENTS = ["When", "Why", "Who", "What", "Where", "How"]

# Custom elements
ELEMENTS = ["Who", "What", "When", "Frequency", "System", "Documentation"]
```

### Retry and Timeout Settings
Adjust for your network conditions:

```python
# Conservative settings (slower but more reliable)
MAX_RETRIES = 10
RETRY_DELAY = 2.0
REQUEST_TIMEOUT = 180.0

# Aggressive settings (faster but may fail more)
MAX_RETRIES = 3
RETRY_DELAY = 0.5
REQUEST_TIMEOUT = 60.0
```

## 🔍 Troubleshooting

### Common Issues

#### ❌ "Authentication failed"
```
Solution: Check your API key
- Ensure OPENAI_API_KEY is set correctly
- Verify the key starts with 'sk-'
- Check if the key has expired
```

#### ❌ "Missing required columns"
```
Solution: Check your Excel headers
- Ensure exact column names (case-sensitive)
- No extra spaces in column names
- All required columns present
```

#### ❌ "Rate limit exceeded"
```
Solution: Adjust retry settings
- Increase RETRY_DELAY
- Increase MAX_RETRY_DELAY
- Use a lower tier model (gpt-3.5-turbo)
```

#### ❌ "File not found"
```
Solution: Check file paths
- Ensure INPUT_FILE exists
- Use absolute paths if needed
- Check file permissions
```

## 📈 Performance Optimization

### For Large Datasets (100+ controls)
```python
# Optimize for bulk processing
MODEL = "gpt-3.5-turbo"  # Faster and cheaper
MAX_RETRIES = 3
RETRY_DELAY = 1.0
```

### For Maximum Accuracy
```python
# Optimize for quality
MODEL = "gpt-4"
MAX_RETRIES = 5
RETRY_DELAY = 2.0
```

## 🛠️ Custom Prompts

You can modify the analysis prompts in the respective functions:

### Control Completeness Analysis
Edit the `ask_llm()` function to customize how controls are evaluated.

### Risk Assessment
Modify `ask_control_objective()` to change risk mitigation analysis.

### Execution Analysis
Update `ask_execution_appropriateness()` for custom automation assessment.

## 🔒 Security Best Practices

1. **Never commit API keys** to version control
2. **Use environment variables** for sensitive data
3. **Rotate API keys** regularly
4. **Monitor API usage** to detect unauthorized access
5. **Use least privilege** API keys when possible

## 💡 Tips for Best Results

### Input Data Quality
- ✅ Write detailed control descriptions (50+ words)
- ✅ Include specific systems, roles, and processes
- ✅ Mention frequency and documentation requirements
- ✅ Describe the control's purpose and methodology

### Model Selection
- Use **GPT-4** for complex controls requiring nuanced analysis
- Use **GPT-3.5-turbo** for simple, well-documented controls
- Use **GPT-4o** for a balance of speed and accuracy

### Batch Processing
- Process controls in batches of 50-100 for optimal performance
- Monitor API costs during large runs
- Save intermediate results to prevent data loss

---

Need help? Check the [GitHub Issues](https://github.com/Mehulchhabra07/Automated-Test-of-Design---Internal-Controls/issues) or contact the maintainer!
