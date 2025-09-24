import pandas as pd
import requests
import json
import os
from typing import Dict, Any
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import docx.oxml
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# File paths
excel_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/leo-matched_snp_genotypes_results666_reordered.xlsx"
existing_doc_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/first_row_report.docx"

def process_drug_names(drug_string):
    """Process drug names to add numbering and proper capitalization"""
    if not drug_string or pd.isna(drug_string):
        return ""
    
    # Split by semicolon and clean up
    drugs = [drug.strip() for drug in str(drug_string).split(';') if drug.strip()]
    
    if not drugs:
        return ""
    
    # Process each drug: capitalize first letter and add numbering
    processed_drugs = []
    for i, drug in enumerate(drugs, 1):
        # Capitalize first letter and make rest lowercase
        capitalized_drug = drug.capitalize()
        processed_drugs.append(f"{i}. {capitalized_drug}")
    
    return ", ".join(processed_drugs)

def read_first_row_data():
    """Read the first row data from the Excel file"""
    try:
        df = pd.read_excel(excel_path)
        if len(df) > 0:
            first_row = df.iloc[0]
            print(f"✅ Successfully read first row data")
            print(f"   Phenotype: {first_row['Phenotype(s)']}")
            print(f"   Drug: {first_row['Drug(s)']}")
            print(f"   Gene: {first_row['Gene']}")
            print(f"   Variant: {first_row['Variant/Haplotypes']}")
            print(f"   Genotype: {first_row['Genotype/Allele']}")
            return first_row
        else:
            print("❌ Excel file is empty")
            return None
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def read_first_five_rows_data():
    """Read the first five rows data from the Excel file"""
    try:
        df = pd.read_excel(excel_path)
        if len(df) >= 5:
            first_five_rows = df.head(5)
            print(f"✅ Successfully read first 5 rows data")
            for i, row in first_five_rows.iterrows():
                print(f"   Row {i+1}: Phenotype: {row['Phenotype(s)']}, Drug: {row['Drug(s)']}")
            return first_five_rows
        else:
            print(f"✅ Successfully read all {len(df)} rows data (less than 5)")
            for i, row in df.iterrows():
                print(f"   Row {i+1}: Phenotype: {row['Phenotype(s)']}, Drug: {row['Drug(s)']}")
            return df
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def read_all_rows_data():
    """Read all rows data from the Excel file"""
    try:
        df = pd.read_excel(excel_path)
        if len(df) > 0:
            print(f"✅ Successfully read all {len(df)} rows data")
            for i, row in df.iterrows():
                print(f"   Row {i+1}: Phenotype: {row['Phenotype(s)']}, Drug: {row['Drug(s)']}")
            return df
        else:
            print("❌ Excel file is empty")
            return None
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def make_openapi_call(drug_name: str) -> Dict[Any, Any]:
    """
    Make OpenAPI call with the detailed pharmacology prompt
    """
    
    # Process drug names for formatting
    processed_drug_names = process_drug_names(drug_name)
    
    # The simplified prompt template for JSON-only response
    prompt = f"""
You are a highly knowledgeable pharmacology assistant specialized in clinical pharmacokinetics, pharmacogenomics, and evidence-based therapeutics.

Please provide comprehensive drug information for **{drug_name}** in the following JSON format only. Do not include any other text or explanations outside the JSON structure.

Provide evidence-based, up-to-date information suitable for healthcare professionals, including:
- Drug overview, mechanism of action, approved indications
- Pharmacokinetic parameters with sources
- Interindividual variability and drug interactions
- Therapeutic drug monitoring recommendations
- Pharmacogenomic considerations (CYP2D6, CYP3A4, etc.)
- Personalized dosing recommendations
- High-quality references from official sources (FDA, CPIC, DPWG, AGNP guidelines)

Return ONLY the JSON response in this exact format:

```json
{{
  "content": [
    {{
      "type": "paragraph",
      "order": 1,
      "title": "1. Therapeutic Drug Monitoring (TDM)",
      "text": "TDM recommendations, therapeutic plasma concentration ranges, sampling time, and interpretation guidance for {drug_name}"
    }},
    {{
      "type": "paragraph",
      "order": 2,
      "title": "2. Personalized Dosing Recommendations",
      "text": "Initial/maintenance dosing, titration schedule, tapering guidance, monitoring parameters, and actionable adjustments for {drug_name}"
    }},
    {{
      "type": "paragraph",
      "order": 3,
      "title": "3. Pharmacogenomics",
      "text": "CYP2D6, CYP3A4, and other relevant genetic polymorphisms, genotype-based recommendations for {drug_name}"
    }},
    {{
      "type": "table",
      "order": 4,
      "title": "Table 1: Pharmacokinetic Parameters ({processed_drug_names})",
      "headers": ["Parameter", "Value / Range", "Source"],
      "rows": [
        ["Oral Bioavailability", "Value with units", "Reference"],
        ["Tmax", "Value with units", "Reference"],
        ["Volume of Distribution (Vd)", "Value with units", "Reference"],
        ["Plasma Protein Binding", "Value with units", "Reference"],
        ["Elimination Half-life (t½)", "Value with units", "Reference"],
        ["Metabolizing Enzymes", "Enzymes involved", "Reference"],
        ["Active/Inactive Metabolites", "Metabolite details", "Reference"],
        ["Clearance Route", "Hepatic/renal percentages", "Reference"],
        ["Time to Steady State", "Value with units", "Reference"]
      ]
    }},
    {{
      "type": "references",
      "order": 5,
      "title": "References",
      "entries": [
        {{ "id": "[1]", "text": "FDA Label for {processed_drug_names}", "link": "https://www.accessdata.fda.gov/scripts/cder/drugsatfda/index.cfm" }},
        {{ "id": "[2]", "text": "CPIC Guidelines for {processed_drug_names}", "link": "https://cpicpgx.org/guidelines/" }},
        {{ "id": "[3]", "text": "AGNP TDM Guidelines", "link": "https://www.agnp.de/page/tdm-guidelines" }},
        {{ "id": "[4]", "text": "DrugBank Database", "link": "https://go.drugbank.com/" }},
        {{ "id": "[5]", "text": "Micromedex Database", "link": "https://www.micromedexsolutions.com/" }},
        {{ "id": "[6]", "text": "Lexicomp Database", "link": "https://www.lexicomp.com/" }}
      ]
    }}
  ]
}}
```

Provide ONLY the JSON response with comprehensive, evidence-based information for {drug_name}.
"""
    
    # API configuration
    api_url = "https://api.openai.com/v1/chat/completions"
    api_key = "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"
    
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json"
    }
    
    payload = {
        "model": "gpt-4o",  # GPT-4.1 (gpt-4o is the latest version)
        "messages": [
            {
                "role": "system",
                "content": "You are a highly knowledgeable pharmacology assistant specialized in clinical pharmacokinetics, pharmacogenomics, and evidence-based therapeutics. Provide comprehensive, accurate, and up-to-date drug information following the specified format exactly."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 4000,  # Increased for comprehensive response
        "temperature": 0.2   # Lower temperature for more consistent, factual responses
    }
    
    try:
        print(f"🔍 Making API call for drug: {drug_name}")
        response = requests.post(api_url, headers=headers, json=payload, timeout=60)
        
        if response.status_code == 200:
            result = response.json()
            print("✅ API call successful")
            return result
        else:
            print(f"❌ API call failed with status code: {response.status_code}")
            print(f"   Response: {response.text}")
            return None
            
    except requests.exceptions.RequestException as e:
        print(f"❌ Error making API call: {e}")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ Error parsing API response: {e}")
        return None

def save_response_to_file(response_data: Dict[Any, Any], drug_name: str):
    """Save the API response to a JSON file"""
    try:
        # Clean drug name for filename
        clean_drug_name = drug_name.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '_')
        filename = f"drug_info_{clean_drug_name}.json"
        
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(response_data, f, indent=2, ensure_ascii=False)
        print(f"✅ Response saved to: {filename}")
        return filename
    except Exception as e:
        print(f"❌ Error saving response: {e}")
        return None

def extract_text_content(response_data: Dict[Any, Any]) -> str:
    """Extract the text content from the API response"""
    try:
        if 'choices' in response_data and len(response_data['choices']) > 0:
            content = response_data['choices'][0]['message']['content']
            
            # Try to extract JSON from the content if it's wrapped in markdown code blocks
            if '```json' in content and '```' in content:
                # Extract JSON from markdown code blocks
                start_idx = content.find('```json') + 7
                end_idx = content.find('```', start_idx)
                if end_idx != -1:
                    json_content = content[start_idx:end_idx].strip()
                    return json_content
            
            # If no markdown blocks, try to find JSON object
            if content.strip().startswith('{') and content.strip().endswith('}'):
                return content.strip()
            
            # If still no JSON found, return the raw content
            return content
        else:
            return "No content found in response"
    except Exception as e:
        print(f"❌ Error extracting content: {e}")
        return "Error extracting content"

def save_text_content(content: str, drug_name: str):
    """Save the text content to a markdown file"""
    try:
        clean_drug_name = drug_name.replace(' ', '_').replace('/', '_').replace('\\', '_').replace(':', '_')
        filename = f"drug_info_{clean_drug_name}.md"
        
        with open(filename, 'w', encoding='utf-8') as f:
            f.write(content)
        print(f"✅ Text content saved to: {filename}")
        return filename
    except Exception as e:
        print(f"❌ Error saving text content: {e}")
        return None

def insert_paragraph_before(ref_para, text="", style=None):
    """
    Inserts a new paragraph with `text` immediately before `ref_para`.
    Returns the new Paragraph object.
    """
    # Use the standard insert_paragraph_before method
    new_para = ref_para.insert_paragraph_before(text)
    
    # Apply style if specified
    if style:
        new_para.style = style
    
    return new_para

def insert_drug_info_into_row_specific_section(doc, row_number: int, content: str, drug_name: str):
    """Insert drug information into a specific row's section of the document"""
    try:
        # Find the phenotype section for this specific row number
        # We'll look for the phenotype section that corresponds to this row
        phenotype_index = -1
        current_row = 0
        
        for i, paragraph in enumerate(doc.paragraphs):
            # Look for phenotype titles (usually main headings)
            if paragraph.style.name.startswith('Heading 1') or paragraph.style.name.startswith('Title'):
                current_row += 1
                if current_row == row_number:
                    phenotype_index = i
                    break
        
        if phenotype_index == -1:
            print(f"❌ Could not find phenotype section for row {row_number}")
            return False
        
        print(f"✅ Found phenotype section for row {row_number} at paragraph {phenotype_index}")
        
        # Parse JSON content
        try:
            json_data = json.loads(content)
            print(f"✅ Successfully parsed JSON data with {len(json_data.get('content', []))} items")
            
            # Find the insertion point (before "How to use the test results" in this phenotype section)
            insertion_index = -1
            for i in range(phenotype_index, len(doc.paragraphs)):
                if "How to use the test results" in doc.paragraphs[i].text:
                    insertion_index = i
                    break
            
            if insertion_index == -1:
                print(f"❌ Could not find 'How to use the test results' section for row {row_number}")
                return False
            
            ref_para = doc.paragraphs[insertion_index]
            
            # Add drug information title
            processed_drug_name = process_drug_names(drug_name)
            title_para = ref_para.insert_paragraph_before(f"Drug Information: {processed_drug_name}")
            title_run = title_para.runs[0]
            title_run.font.name = "Times New Roman"
            title_run.font.size = Pt(14)
            title_run.bold = True
            title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            # Process each content item
            for item in json_data.get("content", []):
                typ = item.get("type")
                title = item.get("title", "")
                text = item.get("text", "")
                
                print(f"Processing item: {typ} - {title}")
                
                if typ == "paragraph":
                    # Remove numbering from title (e.g., "1. Drug Overview" -> "Drug Overview")
                    clean_title = title
                    if title and '. ' in title:
                        clean_title = title.split('. ', 1)[1] if title.split('. ', 1)[1] else title
                    
                    h_para = ref_para.insert_paragraph_before(clean_title)
                    h_run = h_para.runs[0]
                    h_run.font.name = "Times New Roman"
                    h_run.font.size = Pt(12)
                    h_run.bold = True
                    h_run.font.color.rgb = RGBColor(0, 0, 0)
                    
                    if text:
                        t_para = ref_para.insert_paragraph_before(text)
                        t_run = t_para.runs[0]
                        t_run.font.name = "Times New Roman"
                        t_run.font.size = Pt(11)
                        t_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                
                elif typ == "table":
                    # Remove "Table 1:" and "Table 2:" prefixes from table titles
                    clean_title = title
                    if title and ':' in title:
                        # Remove "Table 1:" or "Table 2:" prefix
                        if title.startswith('Table 1:') or title.startswith('Table 2:'):
                            clean_title = title.split(':', 1)[1].strip()
                    
                    # heading
                    tbl_h = ref_para.insert_paragraph_before(clean_title)
                    tbl_h_run = tbl_h.runs[0]
                    tbl_h_run.font.name = "Times New Roman"
                    tbl_h_run.font.size = Pt(12)
                    tbl_h_run.bold = True
                    tbl_h_run.font.color.rgb = RGBColor(0, 0, 0)
                    
                    headers = item["headers"]
                    rows = item["rows"]
                    
                    # build table at doc end (temporarily)
                    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
                    tbl.style = "Table Grid"
                    
                    # headers
                    for ci, hdr in enumerate(headers):
                        cell = tbl.rows[0].cells[ci]
                        cell.text = hdr
                        for run in cell.paragraphs[0].runs:
                            run.font.name = "Times New Roman"
                            run.font.size = Pt(11)
                            run.bold = True
                            run.font.color.rgb = RGBColor(255, 255, 255)
                        cell._tc.get_or_add_tcPr().append(
                            docx.oxml.parse_xml(
                                '<w:shd xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:fill="2F5496"/>'
                            )
                        )
                    
                    # rows
                    for ri, row in enumerate(rows, start=1):
                        for ci, val in enumerate(row):
                            cell = tbl.rows[ri].cells[ci]
                            cell.text = str(val)
                            for run in cell.paragraphs[0].runs:
                                run.font.name = "Times New Roman"
                                run.font.size = Pt(10)
                    
                    # move table before ref_para
                    tbl_elm = tbl._element
                    tbl_elm.getparent().remove(tbl_elm)
                    ref_para._element.addprevious(tbl_elm)
                
                elif typ == "references":
                    # heading
                    ref_h = ref_para.insert_paragraph_before(item.get("title", "References"))
                    ref_h_run = ref_h.runs[0]
                    ref_h_run.font.name = "Times New Roman"
                    ref_h_run.font.size = Pt(12)
                    ref_h_run.bold = True
                    ref_h_run.font.color.rgb = RGBColor(0, 0, 0)
                    
                    for entry in item.get("entries", []):
                        # Include link if available
                        if "link" in entry:
                            text = f"{entry['id']} {entry['text']} - {entry['link']}"
                        else:
                            text = f"{entry['id']} {entry['text']}"
                        
                        rp = ref_para.insert_paragraph_before(text)
                        rp_run = rp.runs[0]
                        rp_run.font.name = "Times New Roman"
                        rp_run.font.size = Pt(10)
                        rp.paragraph_format.left_indent = Pt(20)
            
            return True
        
        except json.JSONDecodeError as e:
            print(f"❌ Error parsing JSON content: {e}")
            return False
        
    except Exception as e:
        print(f"❌ Error inserting drug info into row section: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    print("🔬 Drug Information API Caller")
    print("=" * 50)
    
    # Check if the existing document exists
    if not os.path.exists(existing_doc_path):
        print(f"❌ Existing document not found: {existing_doc_path}")
        return
    
    # Load the existing document once
    doc = Document(existing_doc_path)
    print(f"✅ Loaded existing document: {existing_doc_path}")
    
    # Read first five rows data
    first_five_rows = read_first_five_rows_data()
    if first_five_rows is None:
        return
    
    # Process each row and append to the document
    for index, row in first_five_rows.iterrows():
        drug_name = str(row['Drug(s)'])
        phenotype = str(row['Phenotype(s)'])
        
        print(f"\n💊 Processing row {index + 1}/5: {phenotype} - {drug_name}")
        
        # Make API call
        response = make_openapi_call(drug_name)
        
        if response:
            # Extract and print message content from response
            content = extract_text_content(response)
            if content and content != "No content found in response":
                print(f"✅ API Response Content received for {drug_name}")
                
                # Insert drug information into the phenotype section
                success = insert_drug_info_into_row_specific_section(doc, index + 1, content, drug_name)
                
                if success:
                    print(f"✅ Added drug information for {drug_name} to row {index + 1} section")
                else:
                    print(f"❌ Failed to add drug information for {drug_name} to row {index + 1} section")
            else:
                print(f"❌ No valid content received for {drug_name}")
        else:
            print(f"❌ No response received from API for {drug_name}")
    
    # Save the final document with all drug information
    output_filename = "final_report_with_all_drugs.docx"
    print(f"\n💾 Saving final document as: {output_filename}")
    doc.save(output_filename)
    print(f"✅ Final document saved with all drug information: {output_filename}")
    print(f"\n🎉 Completed processing all 5 phenotypes!")

if __name__ == "__main__":
    main() 