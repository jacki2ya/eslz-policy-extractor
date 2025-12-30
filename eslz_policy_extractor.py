#!/usr/bin/env python3
"""
ESLZ Policy Extractor

Extracts Azure Policy information from the Azure Enterprise Scale Landing Zone
terraform module and generates an Excel spreadsheet with:
- Initiatives worksheet: All policy initiatives assigned via archetypes
- Policies worksheet: All individual policies (directly assigned + expanded from initiatives)

Data flow:
1. GitHub archetype files → assignment names + scope (which archetype)
2. GitHub policy assignment files → assignment details (only for assignments used by archetypes)
3. AzAdvertizer → initiative breakdowns (policies contained in each initiative)
4. AzAdvertizer → ALL policy definition details

Usage:
    python eslz_policy_extractor.py [--output OUTPUT_FILE]
"""

import argparse
import json
import re
import sys
import time
from dataclasses import dataclass, field
from typing import Optional
from urllib.parse import quote

import requests
import xlsxwriter

# Rate limiting configuration
AZADVERTIZER_RATE_LIMIT_SECONDS = 0.2  # 0.2 seconds between AzAdvertizer requests
GITHUB_RATE_LIMIT_SECONDS = 0.1  # Delay between GitHub API requests

# GitHub repository
ESLZ_TF_REPO = "Azure/terraform-azurerm-caf-enterprise-scale"

# AzAdvertizer base URLs
AZADVERTIZER_BASE = "https://www.azadvertizer.net"
AZADVERTIZER_POLICY_HTML = f"{AZADVERTIZER_BASE}/azpolicyadvertizer"
AZADVERTIZER_POLICY_JSON = f"{AZADVERTIZER_BASE}/azpolicyadvertizerjson"
AZADVERTIZER_INITIATIVE_HTML = f"{AZADVERTIZER_BASE}/azpolicyinitiativesadvertizer"
AZADVERTIZER_INITIATIVE_JSON = f"{AZADVERTIZER_BASE}/azpolicyinitiativesadvertizerjson"

# GitHub API
GITHUB_API = "https://api.github.com"
GITHUB_RAW = "https://raw.githubusercontent.com"


def log(msg: str):
    """Print and flush immediately."""
    print(msg)
    sys.stdout.flush()


@dataclass
class PolicyDefinition:
    """Policy definition from AzAdvertizer."""
    name: str
    display_name: str = ""
    description: str = ""
    effect: str = ""
    category: str = ""
    version: str = ""
    policy_type: str = ""  # BuiltIn, Custom
    parameters: str = ""  # Comma-separated list of parameter names
    azadvertizer_url: str = ""


@dataclass
class InitiativeDefinition:
    """Initiative definition from AzAdvertizer."""
    name: str
    display_name: str = ""
    description: str = ""
    category: str = ""
    version: str = ""
    policy_type: str = ""
    policy_count: int = 0
    policy_ids: list = field(default_factory=list)  # List of policy definition IDs
    azadvertizer_url: str = ""


@dataclass
class PolicyAssignment:
    """Policy assignment from GitHub."""
    name: str
    display_name: str = ""
    policy_definition_id: str = ""
    enforcement_mode: str = "Default"
    is_initiative: bool = False
    archetype: str = ""  # Which archetype this assignment belongs to
    github_url: str = ""


class ESLZPolicyExtractor:
    """Extracts policy information from ESLZ and AzAdvertizer."""

    def __init__(self, output_file: str = "eslz_policy_catalog.xlsx"):
        self.output_file = output_file
        self.session = requests.Session()
        self.session.headers.update({
            "User-Agent": "ESLZ-Policy-Extractor/1.0",
            "Accept": "application/json"
        })

        # Data from GitHub
        self.archetype_assignments: dict[str, list[str]] = {}  # archetype -> [assignment_names]
        self.policy_assignments: dict[str, PolicyAssignment] = {}  # assignment_name -> assignment

        # Data from AzAdvertizer
        self.policy_definitions: dict[str, PolicyDefinition] = {}  # policy_id -> definition
        self.initiative_definitions: dict[str, InitiativeDefinition] = {}  # initiative_id -> definition

        # Output data
        self.initiative_rows: list[dict] = []
        self.policy_rows: list[dict] = []

        self._last_azadvertizer_request = 0
        self._last_github_request = 0

    def _rate_limit_azadvertizer(self):
        """Enforce rate limiting for AzAdvertizer requests."""
        elapsed = time.time() - self._last_azadvertizer_request
        if elapsed < AZADVERTIZER_RATE_LIMIT_SECONDS:
            time.sleep(AZADVERTIZER_RATE_LIMIT_SECONDS - elapsed)
        self._last_azadvertizer_request = time.time()

    def _rate_limit_github(self):
        """Enforce rate limiting for GitHub requests."""
        elapsed = time.time() - self._last_github_request
        if elapsed < GITHUB_RATE_LIMIT_SECONDS:
            time.sleep(GITHUB_RATE_LIMIT_SECONDS - elapsed)
        self._last_github_request = time.time()

    def _fetch_json(self, url: str, rate_limit_func=None) -> Optional[dict]:
        """Fetch JSON from URL."""
        if rate_limit_func:
            rate_limit_func()
        try:
            response = self.session.get(url, timeout=30)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            return response.json()
        except (requests.RequestException, json.JSONDecodeError):
            return None

    def _fetch_text(self, url: str, rate_limit_func=None) -> Optional[str]:
        """Fetch text from URL."""
        if rate_limit_func:
            rate_limit_func()
        try:
            response = self.session.get(url, timeout=30)
            if response.status_code == 404:
                return None
            response.raise_for_status()
            return response.text
        except requests.RequestException:
            return None

    def _extract_id_from_path(self, path: str) -> str:
        """Extract policy/initiative ID from ARM resource path."""
        if not path:
            return ""
        return path.rstrip("/").split("/")[-1]

    def _is_initiative(self, policy_definition_id: str) -> bool:
        """Check if ID refers to an initiative (policySetDefinition)."""
        return "policySetDefinitions" in policy_definition_id if policy_definition_id else False

    def _is_uuid(self, s: str) -> bool:
        """Check if string is a UUID format."""
        uuid_pattern = r'^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$'
        return bool(re.match(uuid_pattern, s))

    # =========================================================================
    # STEP 1: Fetch archetype definitions from GitHub
    # =========================================================================
    def fetch_archetypes(self):
        """Fetch archetype definitions to get assignment names and scopes."""
        log("\n[1/4] Fetching archetype definitions from GitHub...")

        url = f"{GITHUB_API}/repos/{ESLZ_TF_REPO}/contents/modules/archetypes/lib/archetype_definitions"
        contents = self._fetch_json(url, self._rate_limit_github) or []

        for item in contents:
            if item.get("type") != "file":
                continue
            filename = item.get("name", "")
            if not filename.endswith((".json", ".tmpl.json")):
                continue
            if "default_empty" in filename:
                continue

            raw_url = item.get("download_url", "")
            if not raw_url:
                continue

            content = self._fetch_text(raw_url, self._rate_limit_github)
            if not content:
                continue

            try:
                # Handle template variables
                content = re.sub(r'\$\{[^}]+\}', 'TEMPLATE_VAR', content)
                data = json.loads(content)

                for archetype_id, archetype_data in data.items():
                    assignments = archetype_data.get("policy_assignments", [])
                    if assignments:
                        self.archetype_assignments[archetype_id] = assignments
                        log(f"  {archetype_id}: {len(assignments)} assignments")

            except json.JSONDecodeError:
                continue

        total_assignments = sum(len(a) for a in self.archetype_assignments.values())
        log(f"  Total: {len(self.archetype_assignments)} archetypes, {total_assignments} assignment references")

    # =========================================================================
    # STEP 2: Fetch policy assignment files from GitHub (only those used by archetypes)
    # =========================================================================
    def fetch_assignments(self):
        """Fetch policy assignment details for assignments used by archetypes."""
        log("\n[2/4] Fetching policy assignment details from GitHub...")

        # Get unique assignment names across all archetypes
        all_assignment_names = set()
        for archetype, assignments in self.archetype_assignments.items():
            for assignment_name in assignments:
                all_assignment_names.add(assignment_name)

        log(f"  Unique assignments to fetch: {len(all_assignment_names)}")

        # Build mapping of assignment name to archetype(s)
        assignment_to_archetypes: dict[str, list[str]] = {}
        for archetype, assignments in self.archetype_assignments.items():
            for assignment_name in assignments:
                if assignment_name not in assignment_to_archetypes:
                    assignment_to_archetypes[assignment_name] = []
                assignment_to_archetypes[assignment_name].append(archetype)

        # Fetch the assignment files
        url = f"{GITHUB_API}/repos/{ESLZ_TF_REPO}/contents/modules/archetypes/lib/policy_assignments"
        contents = self._fetch_json(url, self._rate_limit_github) or []

        # Build filename to URL mapping
        file_map = {}
        for item in contents:
            if item.get("type") == "file":
                file_map[item.get("name", "")] = {
                    "download_url": item.get("download_url", ""),
                    "html_url": item.get("html_url", "")
                }

        found = 0
        for assignment_name in all_assignment_names:
            # Try to find the assignment file
            possible_filenames = [
                f"policy_assignment_es_{assignment_name.lower().replace('-', '_')}.tmpl.json",
                f"policy_assignment_es_{assignment_name.lower().replace('-', '_')}.json",
            ]

            file_info = None
            for fn in possible_filenames:
                if fn in file_map:
                    file_info = file_map[fn]
                    break

            if not file_info:
                # Try case-insensitive search
                for fn, info in file_map.items():
                    normalized = fn.replace("policy_assignment_es_", "").replace(".tmpl.json", "").replace(".json", "")
                    normalized = normalized.replace("_", "-")
                    if normalized.lower() == assignment_name.lower():
                        file_info = info
                        break

            if not file_info:
                log(f"    Warning: Could not find file for assignment '{assignment_name}'")
                continue

            content = self._fetch_text(file_info["download_url"], self._rate_limit_github)
            if not content:
                continue

            try:
                content = re.sub(r'\$\{[^}]+\}', 'TEMPLATE_VAR', content)
                data = json.loads(content)

                name = data.get("name", assignment_name)
                properties = data.get("properties", {})
                policy_def_id = properties.get("policyDefinitionId", "")

                # Create assignment for each archetype it belongs to
                archetypes = assignment_to_archetypes.get(assignment_name, [])
                for archetype in archetypes:
                    assignment = PolicyAssignment(
                        name=name,
                        display_name=properties.get("displayName", name),
                        policy_definition_id=policy_def_id,
                        enforcement_mode=properties.get("enforcementMode", "Default"),
                        is_initiative=self._is_initiative(policy_def_id),
                        archetype=archetype,
                        github_url=file_info["html_url"]
                    )
                    # Use archetype+name as key since same assignment can be in multiple archetypes
                    key = f"{archetype}:{name}"
                    self.policy_assignments[key] = assignment
                    found += 1

            except json.JSONDecodeError:
                continue

        log(f"  Loaded {found} assignment instances")

    # =========================================================================
    # STEP 3: Fetch initiative and policy definitions from AzAdvertizer
    # =========================================================================
    def fetch_from_azadvertizer(self):
        """Fetch all initiative and policy definitions from AzAdvertizer."""
        log("\n[3/4] Fetching definitions from AzAdvertizer...")

        # Collect all policy/initiative IDs we need to fetch
        initiative_ids = set()
        policy_ids = set()

        for assignment in self.policy_assignments.values():
            policy_id = self._extract_id_from_path(assignment.policy_definition_id)
            if not policy_id:
                continue
            if assignment.is_initiative:
                initiative_ids.add(policy_id)
            else:
                policy_ids.add(policy_id)

        log(f"  Initiatives to fetch: {len(initiative_ids)}")
        log(f"  Direct policy assignments to fetch: {len(policy_ids)}")

        # Estimate time
        # We'll fetch initiatives first, then extract policy IDs, then fetch those
        # For now, estimate based on initiatives + direct policies
        initial_estimate = (len(initiative_ids) + len(policy_ids)) * AZADVERTIZER_RATE_LIMIT_SECONDS
        log(f"  Initial estimate (before initiative expansion): {initial_estimate:.0f}s ({initial_estimate/60:.1f} min)")

        # Fetch initiatives
        log("  Fetching initiatives...")
        fetch_count = 0
        for initiative_id in initiative_ids:
            fetch_count += 1
            log(f"    [{fetch_count}/{len(initiative_ids)}] {initiative_id}")
            initiative = self._fetch_initiative_from_azadvertizer(initiative_id)
            if initiative:
                self.initiative_definitions[initiative_id] = initiative
                # Add contained policies to the list
                for pid in initiative.policy_ids:
                    extracted = self._extract_id_from_path(pid)
                    if extracted:
                        policy_ids.add(extracted)

        log(f"  Total policies to fetch (after expansion): {len(policy_ids)}")
        total_estimate = len(policy_ids) * AZADVERTIZER_RATE_LIMIT_SECONDS
        log(f"  Estimated time for policy fetches: {total_estimate:.0f}s ({total_estimate/60:.1f} min)")

        # Fetch policies
        log("  Fetching policy definitions...")
        fetch_count = 0
        for policy_id in policy_ids:
            fetch_count += 1
            if fetch_count % 10 == 0 or fetch_count == len(policy_ids):
                log(f"    [{fetch_count}/{len(policy_ids)}] {policy_id}")
            policy = self._fetch_policy_from_azadvertizer(policy_id)
            if policy:
                self.policy_definitions[policy_id] = policy

        log(f"  Fetched {len(self.policy_definitions)} policy definitions")

    def _fetch_initiative_from_azadvertizer(self, initiative_id: str) -> Optional[InitiativeDefinition]:
        """Fetch initiative definition from AzAdvertizer."""
        self._rate_limit_azadvertizer()

        html_url = f"{AZADVERTIZER_INITIATIVE_HTML}/{quote(initiative_id)}.html"
        data = self._extract_definition_from_html(html_url)

        if data:
            properties = data.get("properties", data)
            policy_defs = properties.get("policyDefinitions", [])
            policy_ids = [p.get("policyDefinitionId", "") for p in policy_defs]

            return InitiativeDefinition(
                name=initiative_id,
                display_name=properties.get("displayName", initiative_id),
                description=properties.get("description", ""),
                category=properties.get("metadata", {}).get("category", ""),
                version=properties.get("metadata", {}).get("version", ""),
                policy_type=properties.get("policyType", ""),
                policy_count=len(policy_ids),
                policy_ids=policy_ids,
                azadvertizer_url=html_url
            )

        # Return minimal definition if HTML extraction failed
        return InitiativeDefinition(
            name=initiative_id,
            display_name=initiative_id,
            azadvertizer_url=html_url
        )

    def _extract_definition_from_html(self, html_url: str) -> Optional[dict]:
        """Extract policy/initiative definition from AzAdvertizer HTML page.

        AzAdvertizer embeds the full definition JSON in a copyDef() JavaScript function.
        """
        try:
            response = self.session.get(html_url, timeout=30)
            if response.status_code != 200:
                return None

            html = response.text

            # Look for the copyDef function which contains the JSON
            # Pattern: function copyDef() { const obj = { ... };
            match = re.search(r'function\s+copyDef\s*\(\s*\)\s*\{\s*const\s+obj\s*=\s*(\{[\s\S]*?\});', html)
            if match:
                json_str = match.group(1)
                return json.loads(json_str)

            return None
        except (requests.RequestException, json.JSONDecodeError):
            return None

    def _fetch_policy_from_azadvertizer(self, policy_id: str) -> Optional[PolicyDefinition]:
        """Fetch policy definition from AzAdvertizer."""
        if policy_id in self.policy_definitions:
            return self.policy_definitions[policy_id]

        self._rate_limit_azadvertizer()

        html_url = f"{AZADVERTIZER_POLICY_HTML}/{quote(policy_id)}.html"
        data = self._extract_definition_from_html(html_url)

        if data:
            properties = data.get("properties", data)
            effect = self._extract_effect(properties)
            parameters = self._extract_parameters(properties)

            return PolicyDefinition(
                name=policy_id,
                display_name=properties.get("displayName", policy_id),
                description=properties.get("description", ""),
                effect=effect,
                category=properties.get("metadata", {}).get("category", ""),
                version=properties.get("metadata", {}).get("version", ""),
                policy_type=properties.get("policyType", ""),
                parameters=parameters,
                azadvertizer_url=html_url
            )

        # Return minimal definition if HTML extraction failed
        return PolicyDefinition(
            name=policy_id,
            display_name=policy_id,
            azadvertizer_url=html_url
        )

    def _extract_parameters(self, properties: dict) -> str:
        """Extract parameter names from policy properties."""
        params = properties.get("parameters", {})
        if not params:
            return ""
        # Return comma-separated list of parameter names
        return ", ".join(sorted(params.keys()))

    def _extract_effect(self, properties: dict) -> str:
        """Extract effect from policy properties."""
        rule = properties.get("policyRule", {})
        then_block = rule.get("then", {})
        effect = then_block.get("effect", "")

        if isinstance(effect, str):
            if effect.startswith("[parameters("):
                # Parameterized effect - get default
                match = re.search(r"\[parameters\('([^']+)'\)\]", effect)
                if match:
                    param_name = match.group(1)
                    params = properties.get("parameters", {})
                    param_def = params.get(param_name, params.get(param_name.lower(), {}))
                    return param_def.get("defaultValue", "Parameterized")
            return effect
        return "Unknown"

    # =========================================================================
    # STEP 4: Build output and generate Excel
    # =========================================================================
    def build_output(self):
        """Build the output data structures."""
        log("\n[4/4] Building output...")

        for key, assignment in self.policy_assignments.items():
            policy_id = self._extract_id_from_path(assignment.policy_definition_id)

            if assignment.is_initiative:
                initiative = self.initiative_definitions.get(policy_id)
                if initiative:
                    # Add initiative row
                    self.initiative_rows.append({
                        "assignment_name": assignment.name,
                        "initiative_name": initiative.name,
                        "initiative_display_name": initiative.display_name,
                        "archetype": assignment.archetype,
                        "enforcement_mode": assignment.enforcement_mode,
                        "policy_count": initiative.policy_count,
                        "category": initiative.category,
                        "version": initiative.version,
                        "azadvertizer_url": initiative.azadvertizer_url,
                        "github_url": assignment.github_url,
                    })

                    # Add policy rows for each policy in initiative
                    for pid in initiative.policy_ids:
                        extracted_pid = self._extract_id_from_path(pid)
                        policy = self.policy_definitions.get(extracted_pid)
                        if policy:
                            self.policy_rows.append({
                                "policy_name": policy.name,
                                "policy_display_name": policy.display_name,
                                "effect": policy.effect,
                                "parameters": policy.parameters,
                                "assignment_type": "Via Initiative",
                                "initiative_name": initiative.name,
                                "initiative_display_name": initiative.display_name,
                                "assignment_name": assignment.name,
                                "archetype": assignment.archetype,
                                "enforcement_mode": assignment.enforcement_mode,
                                "category": policy.category,
                                "azadvertizer_url": policy.azadvertizer_url,
                            })
            else:
                # Direct policy assignment
                policy = self.policy_definitions.get(policy_id)
                if policy:
                    self.policy_rows.append({
                        "policy_name": policy.name,
                        "policy_display_name": policy.display_name,
                        "effect": policy.effect,
                        "parameters": policy.parameters,
                        "assignment_type": "Individual",
                        "initiative_name": "",
                        "initiative_display_name": "",
                        "assignment_name": assignment.name,
                        "archetype": assignment.archetype,
                        "enforcement_mode": assignment.enforcement_mode,
                        "category": policy.category,
                        "azadvertizer_url": policy.azadvertizer_url,
                    })

        log(f"  Initiative rows: {len(self.initiative_rows)}")
        log(f"  Policy rows: {len(self.policy_rows)}")

    def generate_excel(self):
        """Generate the Excel workbook using xlsxwriter."""
        log(f"\nGenerating Excel: {self.output_file}")

        wb = xlsxwriter.Workbook(self.output_file)

        # Formats
        header_format = wb.add_format({
            'bold': True,
            'font_color': 'white',
            'bg_color': '#4472C4',
            'align': 'center',
            'valign': 'vcenter',
            'text_wrap': True,
            'border': 1
        })
        cell_format = wb.add_format({
            'valign': 'top',
            'text_wrap': True,
            'border': 1
        })
        title_format = wb.add_format({
            'bold': True,
            'font_size': 14
        })
        bold_format = wb.add_format({'bold': True})

        # === Initiatives Sheet ===
        ws1 = wb.add_worksheet("Initiatives")

        headers1 = [
            "Assignment Name", "Initiative Definition ID", "Initiative Display Name",
            "Archetype (Scope)", "Enforcement Mode", "Policy Count",
            "Category", "Version", "AzAdvertizer Link (Definition)", "GitHub Link (Assignment)", "Include"
        ]

        for col, h in enumerate(headers1):
            ws1.write(0, col, h, header_format)

        # Add Yes/No data validation for Include column (last column, index 10)
        include_col = len(headers1) - 1  # Last column
        ws1.data_validation(1, include_col, len(self.initiative_rows), include_col, {
            'validate': 'list',
            'source': ['Yes', 'No'],
            'input_title': 'Include in Breakdown',
            'input_message': 'Select Yes to include in Policy Breakdown'
        })

        for row_idx, item in enumerate(self.initiative_rows, 1):
            values = [
                item["assignment_name"], item["initiative_name"], item["initiative_display_name"],
                item["archetype"], item["enforcement_mode"], item["policy_count"],
                item["category"], item["version"], item["azadvertizer_url"], item["github_url"],
                "No"  # Include column defaults to No
            ]
            for col, val in enumerate(values):
                ws1.write(row_idx, col, val, cell_format)

        widths1 = [30, 40, 50, 20, 18, 12, 20, 10, 60, 60, 10]
        for col, w in enumerate(widths1):
            ws1.set_column(col, col, w)
        ws1.freeze_panes(1, 1)
        ws1.autofilter(0, 0, len(self.initiative_rows), len(headers1) - 1)

        # === Policies Sheet ===
        ws2 = wb.add_worksheet("Policies")

        headers2 = [
            "Policy Definition ID", "Policy Display Name", "Effect", "Parameters",
            "Assignment Type", "Initiative Definition ID", "Initiative Display Name",
            "Assignment Name", "Archetype (Scope)", "Enforcement Mode", "Category",
            "AzAdvertizer Link (Definition)"
        ]

        for col, h in enumerate(headers2):
            ws2.write(0, col, h, header_format)

        # Sort: Individual first, then by initiative, then by policy name
        sorted_rows = sorted(
            self.policy_rows,
            key=lambda x: (0 if x["assignment_type"] == "Individual" else 1,
                          x["initiative_name"], x["policy_name"])
        )

        for row_idx, item in enumerate(sorted_rows, 1):
            values = [
                item["policy_name"], item["policy_display_name"], item["effect"],
                item["parameters"], item["assignment_type"], item["initiative_name"],
                item["initiative_display_name"], item["assignment_name"],
                item["archetype"], item["enforcement_mode"], item["category"],
                item["azadvertizer_url"]
            ]
            for col, val in enumerate(values):
                ws2.write(row_idx, col, val, cell_format)

        widths2 = [45, 55, 18, 40, 15, 40, 50, 30, 20, 18, 20, 60]
        for col, w in enumerate(widths2):
            ws2.set_column(col, col, w)
        ws2.freeze_panes(1, 0)
        ws2.autofilter(0, 0, len(sorted_rows), len(headers2) - 1)

        # === Policy Data Sheet (hidden, for FILTER reference) ===
        ws_data = wb.add_worksheet("_PolicyData")
        ws_data.hide()

        # Include a composite key (InitID|Archetype) for scope-aware matching
        data_headers = ["InitDisplayName", "InitID", "Archetype", "InitKey", "PolicyDisplayName", "PolicyID", "Effect", "Parameters", "Category"]
        for col, h in enumerate(data_headers):
            ws_data.write(0, col, h)

        data_row = 1
        for item in sorted_rows:
            if item["assignment_type"] == "Via Initiative":
                init_key = f"{item['initiative_name']}|{item['archetype']}"
                ws_data.write(data_row, 0, item["initiative_display_name"])
                ws_data.write(data_row, 1, item["initiative_name"])
                ws_data.write(data_row, 2, item["archetype"])
                ws_data.write(data_row, 3, init_key)  # Composite key for matching
                ws_data.write(data_row, 4, item["policy_display_name"])
                ws_data.write(data_row, 5, item["policy_name"])
                ws_data.write(data_row, 6, item["effect"])
                ws_data.write(data_row, 7, item["parameters"])
                ws_data.write(data_row, 8, item["category"])
                data_row += 1

        num_data_rows = data_row
        num_initiatives = len(self.initiative_rows) + 1

        # === Policy Breakdown Sheet ===
        ws3 = wb.add_worksheet("Policy Breakdown")

        # Instructions
        ws3.merge_range('A1:H1', "Policy Breakdown - Filtered by Selected Initiatives", title_format)
        ws3.write('A3', "Instructions:", bold_format)
        ws3.write('A4', "1. Go to the 'Initiatives' sheet")
        ws3.write('A5', "2. In the 'Include' column (column K), set 'Yes' for initiatives you want to analyze")
        ws3.write('A6', "3. Return to this sheet - policies are filtered by both initiative AND scope")
        ws3.write('A7', "Note: Requires Excel 365 or Excel 2021+ for dynamic arrays")

        # Headers for the filtered results (includes Scope now)
        breakdown_headers = [
            "Initiative Display Name", "Initiative Definition ID", "Archetype (Scope)",
            "Policy Display Name", "Policy Definition ID", "Effect", "Parameters", "Category"
        ]
        for col, h in enumerate(breakdown_headers):
            ws3.write(8, col, h, header_format)

        # FILTER formula using composite key (InitID|Archetype) for scope-aware matching
        # Initiatives sheet: B = InitID, D = Archetype, K = Include, so composite key = B&"|"&D
        filter_formula = (
            f"=IFERROR(FILTER("
            f"CHOOSECOLS('_PolicyData'!A2:I{num_data_rows},1,2,3,5,6,7,8,9),"
            f"ISNUMBER(MATCH('_PolicyData'!D2:D{num_data_rows},"
            f'FILTER(Initiatives!$B$2:$B${num_initiatives}&"|"&Initiatives!$D$2:$D${num_initiatives},Initiatives!$K$2:$K${num_initiatives}="Yes",""),0))'
            f'),"No initiatives selected - set Include to Yes on Initiatives sheet")'
        )

        ws3.write_dynamic_array_formula('A10', filter_formula)

        widths3 = [50, 40, 20, 55, 45, 18, 40, 25]
        for col, w in enumerate(widths3):
            ws3.set_column(col, col, w)
        ws3.freeze_panes(9, 0)

        wb.close()
        log(f"  Saved: {len(self.initiative_rows)} initiatives, {len(self.policy_rows)} policies")
        log(f"  Policy data rows for breakdown: {num_data_rows - 1}")

    def run(self):
        """Run the extraction."""
        log("=" * 70)
        log("ESLZ Policy Extractor")
        log(f"AzAdvertizer rate limit: {AZADVERTIZER_RATE_LIMIT_SECONDS}s")
        log("=" * 70)

        self.fetch_archetypes()
        self.fetch_assignments()
        self.fetch_from_azadvertizer()
        self.build_output()
        self.generate_excel()

        log("\n" + "=" * 70)
        log("Complete!")
        log("=" * 70)


def main():
    parser = argparse.ArgumentParser(description="Extract ESLZ policy catalog")
    parser.add_argument("-o", "--output", default="eslz_policy_catalog.xlsx")
    args = parser.parse_args()

    ESLZPolicyExtractor(output_file=args.output).run()


if __name__ == "__main__":
    main()
