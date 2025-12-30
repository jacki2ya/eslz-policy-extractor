# ESLZ Policy Extractor

Extracts Azure Policy information from the [Azure Enterprise Scale Landing Zone (ESLZ) Terraform module](https://github.com/Azure/terraform-azurerm-caf-enterprise-scale) and generates an Excel spreadsheet for analysis.

## Features

- Fetches all policy assignments from ESLZ archetype definitions
- Expands policy initiatives to show individual policies
- Links to AzAdvertizer for policy/initiative definitions
- Links to GitHub for assignment configurations
- Scope-aware filtering for policy breakdown analysis

## Requirements

- Python 3.8+
- Excel 365 or Excel 2021+ (for dynamic array formulas in Policy Breakdown sheet)

## Installation

```bash
# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install requests xlsxwriter
```

## Usage

```bash
python eslz_policy_extractor.py
```

Options:
- `-o, --output`: Output filename (default: `eslz_policy_catalog.xlsx`)

## Output

The script generates an Excel workbook with four worksheets:

### 1. Assigned Initiatives

Lists all policy initiatives assigned via ESLZ archetypes:

| Column | Description |
|--------|-------------|
| Assignment Name | Name of the policy assignment |
| Initiative Definition ID | ID of the initiative (UUID or ALZ name) |
| Initiative Display Name | Human-readable name |
| Archetype (Scope) | Which ESLZ archetype assigns this |
| Enforcement Mode | Default or DoNotEnforce |
| Policy Count | Number of policies in the initiative |
| Category | Policy category |
| Version | Initiative version |
| AzAdvertizer Link | Link to definition on AzAdvertizer |
| GitHub Link | Link to assignment file on GitHub |
| Include | Set to "Yes" to include in Policy Breakdown |

### 2. Assigned Policies

Lists policies that are directly assigned (not via initiatives):

| Column | Description |
|--------|-------------|
| Assignment Name | Name of the policy assignment |
| Policy Definition ID | Policy definition ID |
| Policy Display Name | Human-readable policy name |
| Archetype (Scope) | Which ESLZ archetype assigns this |
| Enforcement Mode | Default or DoNotEnforce |
| Effect | Policy effect (Audit, Deny, Deploy, etc.) |
| Parameters | Policy parameters as JSON |
| Category | Policy category |
| AzAdvertizer Link | Link to definition on AzAdvertizer |
| GitHub Link | Link to assignment file on GitHub |
| Include | Set to "Yes" to include in Policy Breakdown |

### 3. Initiative Policies

Lists all policies expanded from initiatives (informational, no Include column):

| Column | Description |
|--------|-------------|
| Initiative Definition ID | Parent initiative ID |
| Initiative Display Name | Parent initiative name |
| Assignment Name | Name of the policy assignment |
| Archetype (Scope) | Which ESLZ archetype assigns this |
| Policy Definition ID | Policy definition ID |
| Policy Display Name | Human-readable policy name |
| Effect | Policy effect (Audit, Deny, Deploy, etc.) |
| Parameters | Policy parameters as JSON |
| Category | Policy category |
| AzAdvertizer Link | Link to definition on AzAdvertizer |

### 4. Policy Breakdown

Dynamic worksheet that combines policies from selected initiatives AND selected direct policies:

1. Go to **Assigned Initiatives** sheet and set **Include** to "Yes" for initiatives to analyze
2. Go to **Assigned Policies** sheet and set **Include** to "Yes" for direct policies to analyze
3. Return to **Policy Breakdown** - shows combined results from both selections
4. Filtering is scope-aware (same item at different scopes won't show duplicates)

## Data Sources

- **GitHub**: ESLZ Terraform module archetype definitions and policy assignments
- **AzAdvertizer**: Policy and initiative definitions with full metadata

## Rate Limiting

The script includes a configurable rate limit for AzAdvertizer requests (default: 0.2s between requests) to avoid overwhelming the service.

## License

MIT

---

*This project was co-authored with [Claude Code](https://claude.com/claude-code).*
