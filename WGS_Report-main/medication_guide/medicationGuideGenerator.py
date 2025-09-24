
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generate_report_clone_from_template.py
Deep-clone prototype blocks from the template so the output inherits EXACT formatting
from the placeholder locations.

Usage:
  python generate_report_clone_from_template.py template_cloneable.docx input.xlsx output.docx

Notes:
- Placeholders must be kept in single runs (don't split "{{TOKEN}}" across runs).
- We replace Drug(s) fullwidth semicolons "；" with " & " (no splitting).
- We do not set fonts/alignment in code; everything comes from the cloned blocks.
"""
import sys, os
import pandas as pd
from copy import deepcopy
from docx import Document
import xml.etree.ElementTree as ET

COL_PHENO = "Phenotype(s)"
COL_DRUG  = "Drug(s)"
COL_GENE  = "Gene"
COL_VAR   = "Variant/Haplotypes"
COL_GT    = "Genotype/Allele"
COL_RES   = "Result"
COL_ANN   = "Annotation Text"
COL_PHENO_CAT = "Phenotype Category"
COL_LEVEL_EVIDENCE = "Level of Evidence"

def _norm(df):
    expected = [COL_PHENO, COL_DRUG, COL_GENE, COL_VAR, COL_GT, COL_RES, COL_ANN, COL_PHENO_CAT, COL_LEVEL_EVIDENCE]
    for c in expected:
        if c not in df.columns:
            raise ValueError(f"Missing column: {c}")
    def s(x):
        if pd.isna(x): return ""
        return str(x).strip()
    for c in expected:
        df[c] = df[c].map(s)
    # Handle multiple drugs separated by semicolons
    def format_drug_names(drug_str):
        if pd.isna(drug_str) or drug_str == "":
            return ""
        # Replace both fullwidth and regular semicolons with " & "
        drug_str = drug_str.replace('；', ';').replace(';', ' & ')
        # Add space before each drug name (after &)
        drug_str = drug_str.replace(' &', ' & ')
        # Remove any leading/trailing spaces
        return drug_str.strip()
    
    df[COL_DRUG] = df[COL_DRUG].apply(format_drug_names)
    return df

def _load_drug_database(database_path):
    """Load drug information from XML database"""
    try:
        tree = ET.parse(database_path)
        root = tree.getroot()
        
        # Define the namespace
        namespace = {'db': 'http://www.drugbank.ca'}
        
        drug_db = {}
        
        # Find all drug entries using the namespace
        for drug in root.findall('.//db:drug', namespace):
            name_elem = drug.find('db:name', namespace)
            if name_elem is not None and name_elem.text:
                drug_name = name_elem.text.strip()
                
                # Extract drug information using namespace
                pharm = drug.find('db:pharmacodynamics', namespace)
                half_life = drug.find('db:half-life', namespace)
                toxicity = drug.find('db:toxicity', namespace)
                route_elim = drug.find('db:route-of-elimination', namespace)
                description = drug.find('db:description', namespace)
                
                drug_db[drug_name.lower()] = {
                    'pharmacodynamics': pharm.text.strip() if pharm is not None and pharm.text else 'Information not available',
                    'half_life': half_life.text.strip() if half_life is not None and half_life.text else 'Information not available',
                    'toxicity': toxicity.text.strip() if toxicity is not None and toxicity.text else 'Information not available',
                    'route_of_elimination': route_elim.text.strip() if route_elim is not None and route_elim.text else 'Information not available',
                    'description': description.text.strip() if description is not None and description.text else 'Information not available'
                }
        
        return drug_db
    except Exception as e:
        print(f"Warning: Could not load drug database: {e}")
        return {}

def _clean_text(text):
    """Remove reference numbers and citations from text"""
    if not text or text == 'Information not available':
        return text
    
    import re
    # Remove reference patterns like [A1115], [A1115,A1116], [A31772,L1214], etc.
    # Pattern matches: [Any letter+numbers, optionally followed by comma and more letter+numbers]
    cleaned_text = re.sub(r'\[[A-Z]\d+(?:,[A-Z]\d+)*\]', '', text)
    
    # Remove any extra spaces that might be left after removing references
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    
    # Clean up any trailing commas or periods that might be left
    cleaned_text = re.sub(r'[,\s]+$', '', cleaned_text)
    
    return cleaned_text.strip()

def _clean_csv_summary_text(text):
    """Clean CSV Summary text by removing <ref>...</ref> tags and extra whitespace."""
    if not isinstance(text, str) or text.strip() == "":
        return ""
    import re
    cleaned = re.sub(r'<ref>.*?</ref>', '', text)
    cleaned = re.sub(r'\s+', ' ', cleaned)
    return cleaned.strip().strip(',').strip()

def _load_drugbank_summary_csv(csv_path):
    """Load DrugBank summaries from CSV into a dict: drug_name(lower) -> summary"""
    if not csv_path or not os.path.exists(csv_path):
        print(f"Warning: DrugBank summary CSV not found: {csv_path}")
        return {}
    try:
        df = pd.read_csv(csv_path)
        if 'Drug_Name' not in df.columns or 'Summary' not in df.columns:
            print("Warning: CSV missing required columns 'Drug_Name' or 'Summary'")
            return {}
        summary_map = {}
        for _, row in df.iterrows():
            name = str(row['Drug_Name']).strip() if pd.notna(row['Drug_Name']) else ''
            summary = str(row['Summary']).strip() if pd.notna(row['Summary']) else ''
            if not name or not summary:
                continue
            summary_map[name.lower()] = _clean_csv_summary_text(summary)
        return summary_map
    except Exception as e:
        print(f"Warning: Could not load DrugBank summary CSV: {e}")
        return {}

def _get_drug_summary_from_map(drug_name, summary_map):
    """Get summary for a drug by exact or partial match from preloaded map."""
    if not summary_map or not drug_name:
        return None
    key = drug_name.lower().strip()
    if key in summary_map:
        return summary_map[key]
    for db_name, summary in summary_map.items():
        if key in db_name or db_name in key:
            return summary
    return None

def _get_drug_info(drug_name, drug_db):
    """Get drug information from database"""
    if not drug_db:
        return {
            'pharmacodynamics': 'Information not available',
            'half_life': 'Information not available',
            'toxicity': 'Information not available',
            'route_of_elimination': 'Information not available'
        }
    
    # Try exact match first
    if drug_name.lower() in drug_db:
        drug_info = drug_db[drug_name.lower()]
        # Clean the text to remove references
        return {
            'pharmacodynamics': _clean_text(drug_info['pharmacodynamics']),
            'half_life': _clean_text(drug_info['half_life']),
            'toxicity': _clean_text(drug_info['toxicity']),
            'route_of_elimination': _clean_text(drug_info['route_of_elimination'])
        }
    
    # Try partial matches for drugs with multiple names
    for db_drug_name, info in drug_db.items():
        if drug_name.lower() in db_drug_name or db_drug_name in drug_name.lower():
            # Clean the text to remove references
            return {
                'pharmacodynamics': _clean_text(info['pharmacodynamics']),
                'half_life': _clean_text(info['half_life']),
                'toxicity': _clean_text(info['toxicity']),
                'route_of_elimination': _clean_text(info['route_of_elimination'])
            }
    
    # Return default if no match found
    return {
        'pharmacodynamics': 'Information not available',
        'half_life': 'Information not available',
        'toxicity': 'Information not available',
        'route_of_elimination': 'Information not available',
        'description': 'Information not available'
    }

def _group(df):
    grouped = {}
    for _, row in df.iterrows():
        ph = row[COL_PHENO]; dr = row[COL_DRUG]
        grouped.setdefault(ph, {})
        grouped[ph].setdefault(dr, [])
        grouped[ph][dr].append(row)
    return grouped

def _get_body(document):
    return document.element.body

def _iter_blocks(body):
    # yields block-level elements (p or tbl)
    for el in list(body.iterchildren()):
        yield el

def _block_text(block):
    # concat all text in a block
    texts = block.xpath('.//w:t')
    return ''.join([t.text or '' for t in texts])

def _collect_range(body, start_marker, end_marker):
    """
    Returns (blocks_between, start_idx, end_idx).
    blocks_between EXCLUDES the marker paragraphs themselves.
    """
    children = list(body.iterchildren())
    start_idx = end_idx = None
    for i, el in enumerate(children):
        if el.tag.endswith('p') and start_marker in _block_text(el):
            start_idx = i
            break
    if start_idx is None:
        raise RuntimeError(f"Start marker not found: {start_marker}")
    for j in range(start_idx+1, len(children)):
        if children[j].tag.endswith('p') and end_marker in _block_text(children[j]):
            end_idx = j
            break
    if end_idx is None:
        raise RuntimeError(f"End marker not found: {end_marker}")
    # slice between start+1 and end_idx (exclusive)
    between = children[start_idx+1:end_idx]
    return between, start_idx, end_idx

def _remove_range(body, start_idx, end_idx):
    # remove from start to end inclusive
    children = list(body.iterchildren())
    for idx in range(end_idx, start_idx-1, -1):
        body.remove(children[idx])

def _replace_texts(block, mapping):
    # Replace placeholders inside a block (paragraph/table)
    for t in block.xpath('.//w:t'):
        if t.text:
            text = t.text
            for k, v in mapping.items():
                text = text.replace(k, v)
            t.text = text

def _append_cloned_blocks(body, prototype_blocks, mapping):
    for proto in prototype_blocks:
        clone = deepcopy(proto)
        _replace_texts(clone, mapping)
        body.append(clone)

def generate(template_path, excel_path, output_path, database_path=None, summary_csv_path=None):
    if not os.path.exists(template_path): raise FileNotFoundError(template_path)
    if not os.path.exists(excel_path): raise FileNotFoundError(excel_path)

    # Load drug database if provided
    drug_db = {}
    if database_path and os.path.exists(database_path):
        drug_db = _load_drug_database(database_path)
        print(f"Loaded drug database with {len(drug_db)} drugs")
    else:
        print("Warning: Drug database not found, using default values")
    # Load DrugBank summaries CSV for drug descriptions
    drug_summaries = _load_drugbank_summary_csv(summary_csv_path) if summary_csv_path else {}
    if drug_summaries:
        print(f"Loaded DrugBank summaries for {len(drug_summaries)} drugs")

    doc = Document(template_path)
    body = _get_body(doc)

    # Collect prototypes
    pheno_proto, s1, e1 = _collect_range(body, "{{PHENOTYPE_BLOCK_START}}", "{{PHENOTYPE_BLOCK_END}}")
    drug_proto,  s2, e2 = _collect_range(body, "{{DRUG_BLOCK_START}}",      "{{DRUG_BLOCK_END}}")
    info_proto,  s3, e3 = _collect_range(body, "{{INFO_GROUP_START}}",      "{{INFO_GROUP_END}}")
    drug_desc_proto, s4, e4 = _collect_range(body, "{{DRUG_DESCRIPTION_BLOCK_START}}", "{{DRUG_DESCRIPTION_BLOCK_END}}")

    # Remove prototype ranges (in reverse order of appearance to keep indexes valid)
    for s, e in sorted([(s1,e1),(s2,e2),(s3,e3),(s4,e4)], key=lambda x: x[0], reverse=True):
        _remove_range(body, s, e)

    # Load data
    df = pd.read_excel(excel_path, engine='openpyxl')
    df = _norm(df)
    grouped = _group(df)

    # Build output by cloning prototypes
    for pheno, drug_map in grouped.items():
        _append_cloned_blocks(
            body, pheno_proto,
            {"{{PHENOTYPE_TITLE}}": pheno, "{{PHENOTYPE_DESC}}": "", "{{PHENOTYPE_KEYWORDS}}": ""}
        )
        for drug, rows in drug_map.items():
            # Drug description: prefer DrugBank CSV Summary; fallback to XML description
            individual_drugs = [d.strip() for d in drug.split('&')]
            descriptions = []
            for dn in individual_drugs:
                summ = _get_drug_summary_from_map(dn, drug_summaries)
                if summ:
                    descriptions.append(summ)
            if descriptions:
                # De-duplicate while preserving order
                seen = set()
                ordered_unique = []
                for s in descriptions:
                    if s not in seen:
                        seen.add(s)
                        ordered_unique.append(s)
                drug_description = ' | '.join(ordered_unique)
            else:
                if individual_drugs:
                    first_drug = individual_drugs[0].strip()
                    drug_info = _get_drug_info(first_drug, drug_db)
                    drug_description = drug_info.get('description', 'Information not available')
                else:
                    drug_description = "Information not available"
            
            _append_cloned_blocks(
                body, drug_proto,
                {"{{DRUG_TITLE}}": drug, "{{DRUG_DESC}}": drug_description, "{{DRUG_KEYWORDS}}": ""}
            )
            for r in rows:
                mapping = {
                    "{{RESULT}}": r[COL_RES],
                    "{{GENE}}": r[COL_GENE],
                    "{{VARIANT}}": r[COL_VAR],
                    "{{GENOTYPE}}": r[COL_GT],
                    "{{ANNOTATION}}": r[COL_ANN],
                    "{{PHENOTYPE_CATEGORY}}": r[COL_PHENO_CAT],
                    "{{LEVEL_OF_EVIDENCE}}": r[COL_LEVEL_EVIDENCE],
                    # Handle the new text section placeholders
                    "{{variant}}": r[COL_VAR],
                    "{{Gene}}": r[COL_GENE],
                    "{{Genotype}}": r[COL_GT],
                    # Handle the new Drug placeholder
                    "{{Drug}}": drug,
                }
                _append_cloned_blocks(body, info_proto, mapping)
            
            # Add drug description block after each drug with information from database
            # Extract individual drug names from the combined drug string
            individual_drugs = [d.strip() for d in drug.split('&')]
            
            # Get drug information for the first drug (or combine if multiple)
            if individual_drugs:
                first_drug = individual_drugs[0].strip()
                drug_info = _get_drug_info(first_drug, drug_db)
                
                drug_desc_mapping = {
                    "{{Pharmacodynamics}}": drug_info['pharmacodynamics'],
                    "{{Route_of_elimination}}": drug_info['route_of_elimination'],
                    "{{Half_life}}": drug_info['half_life'],
                    "{{Toxicity}}": drug_info['toxicity']
                }
            else:
                drug_desc_mapping = {
                    "{{Pharmacodynamics}}": "Information not available",
                    "{{Route_of_elimination}}": "Information not available",
                    "{{Half_life}}": "Information not available",
                    "{{Toxicity}}": "Information not available"
                }
            
            _append_cloned_blocks(body, drug_desc_proto, drug_desc_mapping)

    doc.save(output_path)
    return output_path

def main():
    # Hardcoded file paths
    template_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/medication_guide/template.docx"
    excel_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/Data/preprocessed_data_v1.xlsx"
    database_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/Data/full database.xml"
    summary_csv_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/drug_bank/drug_bank_results.csv"
    output_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/medication_guide/medication_guide.docx"
    
    # Check if files exist
    if not os.path.exists(template_path):
        print(f"Error: Template file not found: {template_path}")
        sys.exit(1)
    
    if not os.path.exists(excel_path):
        print(f"Error: Excel file not found: {excel_path}")
        sys.exit(1)
    
    if not os.path.exists(database_path):
        print(f"Warning: Drug database not found: {database_path}")
        database_path = None
    
    print(f"Template: {template_path}")
    print(f"Input Excel: {excel_path}")
    print(f"Drug Database: {database_path}")
    print(f"DrugBank Summary CSV: {summary_csv_path}")
    print(f"Output: {output_path}")
    
    try:
        out = generate(template_path, excel_path, output_path, database_path, summary_csv_path)
        print(f"Successfully generated: {out}")
    except Exception as e:
        print(f"Error generating report: {e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
