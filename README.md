# Private-Credit-Fund-Risk-Analysis
# PCO Exposure Analysis Tool

A Python-based tool for analyzing Private Credit Opportunities (PCO) fund exposure across multiple dimensions, generating comprehensive Excel reports with charts and formatted outputs.

![Excel Report Example](https://via.placeholder.com/800x400.png?text=Sample+Excel+Output)

## Features

- **Automated Data Processing**
  - Processes RISK tab data from NAV summary files
  - Creates 4 new dimensions using issuer mapping:
    - Issuer Name-N
    - Moody's Industry-N
    - Lien-N
    - Regional-N
- **Multi-Dimensional Analysis**
  - Issuer exposure breakdown
  - Lien type distribution
  - Regional exposure
  - Industry sector analysis
- **Automated Reporting**
  - Excel report generation with:
  - Formatted tables with currency/percentage formatting
  - Interactive pie charts
  - Consolidated PCO/SMA views
- **Smart Configuration**
  - Flexible fund configuration
  - Custom exclusion lists
  - Fixed total commitment tracking ($1.67B)

## Installation

1. **Prerequisites**
   - Python 3.8+
   - pandas
   - openpyxl
   - numpy

2. **Setup**
```bash
git clone https://github.com/yourusername/pco-analysis.git
cd pco-analysis
pip install -r requirements.txt
