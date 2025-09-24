import pandas as pd
from docxtpl import DocxTemplate
from datetime import datetime
import os

# Import drug bank script functionality
from drug_bank_script import fetch_drug_data

# Import Wikipedia script functionality
from wikipediaScript import fetch_phenotype_description, fetch_drug_description

# File paths
excel_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/leo-matched_snp_genotypes_results666_reordered.xlsx"
template_path = "genetic_report_template.docx"
output_path = "final_report.docx"

def preserve_manual_changes():
    """Preserve manual changes by creating a backup before any modifications"""
    if os.path.exists(template_path):
        backup_path = f"genetic_report_template_manual_{datetime.now().strftime('%Y%m%d_%H%M%S')}.docx"
        import shutil
        shutil.copy(template_path, backup_path)
        print(f"✅ Manual changes backed up to: {backup_path}")
        return backup_path
    return None

def restore_template_from_backup(backup_path):
    """Restore template from backup"""
    if os.path.exists(backup_path):
        import shutil
        shutil.copy(backup_path, template_path)
        print(f"✅ Template restored from: {backup_path}")
        return True
    return False

def create_backup_template():
    """Create a backup template with a different name"""
    backup_path = "genetic_report_template_backup.docx"
    create_template()
    import shutil
    shutil.copy(template_path, backup_path)
    print(f"✅ Backup template created: {backup_path}")

def create_template():
    """Create a Word template with Jinja2 placeholders"""
    from docx import Document
    from docx.shared import Cm, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import parse_xml, OxmlElement
    from docx.oxml.ns import nsdecls, qn
    
    doc = Document()
    
    # Add title - using Phenotype as main header
    title = doc.add_heading('{{Phenotype}}', level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Make the title Times New Roman, 14pt, bold, and black color
    for run in title.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(14)  # 14pt font size
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add the three new fields
    doc.add_paragraph()
    
    # Add Wikipedia Info description with Times New Roman, 11pt
    wiki_info_para = doc.add_paragraph('Description：{{Wikipedia_Info}}')
    for run in wiki_info_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    # Add Drug name with Times New Roman, 11pt, and bold
    drug_para = doc.add_paragraph('Drug: {{Drug}}')
    for run in drug_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
        run.bold = True
    
    # Add Wikipedia Additional description with Times New Roman, 11pt
    wiki_add_para = doc.add_paragraph('Description：{{Wikipedia_Additional}}')
    for run in wiki_add_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    # Add rectangular box with genetic guidance text
    box_paragraph = doc.add_paragraph('Genetic Guidance for {{Drug}}\n{{Result}}')
    
    # Format the text with Times New Roman, 11pt, and make only "Genetic Guidance for {{Drug}}" bold
    for run in box_paragraph.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
        # Only make "Genetic Guidance for {{Drug}}" bold, not the Result part
        if 'Genetic Guidance for' in run.text:
            run.bold = True
        else:
            run.bold = False
            # Set color for Result text - red by default, green for "✅ Administer normal dose"
            if '✅ Administer normal dose' in run.text:
                run.font.color.rgb = RGBColor(0, 128, 0)  # Green color
            else:
                run.font.color.rgb = RGBColor(255, 0, 0)  # Red color
    
    # Create a border around the paragraph to make it look like a box
    
    # Add border to the paragraph with same color as table headers
    p = box_paragraph._element
    pPr = p.get_or_add_pPr()
    pBdr = OxmlElement('w:pBdr')
    pBdr.append(parse_xml(f'<w:top {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="2F5496"/>'))
    pBdr.append(parse_xml(f'<w:left {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="2F5496"/>'))
    pBdr.append(parse_xml(f'<w:bottom {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="2F5496"/>'))
    pBdr.append(parse_xml(f'<w:right {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="2F5496"/>'))
    pPr.append(pBdr)
    
    # Add some padding by setting paragraph spacing and center the text
    box_paragraph.paragraph_format.space_before = Cm(0.2)
    box_paragraph.paragraph_format.space_after = Cm(0.2)
    box_paragraph.paragraph_format.right_indent = Cm(0.2)
    box_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Add the genetic analysis paragraph
    genetic_analysis_para = doc.add_paragraph('We have analyzed the {{variant}} locus of your {{Gene}} gene. Your genotype is determined to be {{Genotype}}. {{Annotation_Text}}.')
    
    # Format with Times New Roman, 11pt
    for run in genetic_analysis_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    # Add centered title after the genetic analysis
    subtitle = doc.add_heading('My Genetic Result Details', level=2)
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Make the subtitle Times New Roman, 14pt, bold, and black color
    for run in subtitle.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add table
    table = doc.add_table(rows=2, cols=5)
    table.style = 'Table Grid'
    
    # Header row
    header_cells = table.rows[0].cells
    header_cells[0].text = 'Gene'
    header_cells[1].text = 'Variant'
    header_cells[2].text = 'Genotype'
    header_cells[3].text = 'Phenotype Category'
    header_cells[4].text = 'Level of Evidence'
    
    # Make header cells use the same color as "My Genetic Result Details" heading
    for cell in header_cells:
        # Set cell background to match the heading color (using a dark blue/gray)
        cell._tc.get_or_add_tcPr().append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="2F5496"/>'))
        
        # Set text color to white for contrast and format with Times New Roman, 11pt
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
                run.font.size = Pt(11)
                run.font.color.rgb = RGBColor(255, 255, 255)  # White color
    
    # Data row
    data_cells = table.rows[1].cells
    data_cells[0].text = '{{Gene}}'
    data_cells[1].text = '{{Variant_Haplotypes}}'
    data_cells[2].text = '{{Genotype_Allele}}'
    data_cells[3].text = '{{Phenotype_Category}}'
    data_cells[4].text = '{{Level_of_Evidence}}'
    
    # Format data cells with Times New Roman, 11pt
    for cell in data_cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Times New Roman'
                run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
                run.font.size = Pt(11)
    
    # Add small text after table
    footnote = doc.add_paragraph('* Level of Evidence：1A：Highest reliability；1B: Strong evidence；2A: Moderate reliability with enhanced pharmacogenetic focus ; 2B: Moderate evidence; 3: Emerging evidence; 4: Minimal to no evidence.')
    
    # Format with Times New Roman, 9pt, bold, and black color
    for run in footnote.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(9)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    

    
    # Add bold title "How to use the test results"
    how_to_use_title = doc.add_heading('How to use the test results', level=2)
    how_to_use_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Make the title Times New Roman, 14pt, bold, and black color
    for run in how_to_use_title.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add "Notification" title
    notification_title = doc.add_heading('Notification', level=2)
    notification_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Make the notification title Times New Roman, 12pt, bold, and black color
    for run in notification_title.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(12)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add notification paragraph
    notification_para = doc.add_paragraph('The test does not yet have the qualifications for clinical diagnosis. The test results and medication recommendations provided are for reference only and cannot be directly used as the basis for clinical medication. If you have relevant medication needs, you can consult authoritative clinicians or professionals based on the results of this genetic test to obtain professional medication recommendations or follow the doctor\'s advice for clinically qualified inspection and test.')
    
    # Format with Times New Roman, 11pt
    for run in notification_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    # Add "Test content" title
    test_content_title = doc.add_heading('Test content', level=2)
    test_content_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Make the test content title Times New Roman, 12pt, bold, and black color
    for run in test_content_title.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(12)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add test content paragraph
    test_content_para = doc.add_paragraph('Different genetic test versions and data sources may have different coverage of the above variants.')
    
    # Format with Times New Roman, 11pt
    for run in test_content_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    # Add "Limit of detection" title
    limit_title = doc.add_heading('Limit of detection', level=2)
    limit_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Make the limit title Times New Roman, 12pt, bold, and black color
    for run in limit_title.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(12)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)  # Black color
    
    # Add limit of detection paragraph
    limit_para = doc.add_paragraph('Test results are limited by the current technology and the level of scientific cognition. Depends on product package purchased, the test report may not cover all genes or loci related in each drug response. The accuracy of existing scientific research and the inherent detection error rate associated with the type of detection you purchase may also affect the accuracy of the interpretation of the item.')
    
    # Format with Times New Roman, 11pt
    for run in limit_para.runs:
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
        run.font.size = Pt(11)
    
    doc.save(template_path)
    print(f"✅ Template created with Phenotype and three new fields: {template_path}")

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

def read_excel_data():
    """Read first five rows from the Excel file"""
    try:
        df = pd.read_excel(excel_path)
        if len(df) >= 5:
            first_five_rows = df.head(5)
            print(f"✅ Successfully read first 5 rows from Excel file: {excel_path}")
            print(f"   Total rows in file: {len(df)}")
            print(f"   Processing first 5 rows")
            print(f"   Columns: {list(df.columns)}")
            return first_five_rows
        else:
            print(f"✅ Successfully read all {len(df)} rows from Excel file: {excel_path}")
            print(f"   Columns: {list(df.columns)}")
            return df
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def generate_report(df):
    """Generate the report using Jinja2 template with all row data"""
    if df is None:
        print("❌ No data to process")
        return
    
    # Create template if it doesn't exist
    # if not os.path.exists(template_path):
    #     create_template()
    
    # Check if template exists
    if not os.path.exists(template_path):
        print(f"❌ Template not found: {template_path}")
        print("   Please ensure the template file exists or run create_template() manually")
        return
    
    # Process each row and create separate files, then combine
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
    import docx.oxml
    import tempfile
    
    temp_files = []
    
    # Process each row
    for index, row in df.iterrows():
        print(f"📄 Processing row {index + 1}/{len(df)}: {row['Phenotype(s)']}")
        
        # Prepare data for template - phenotype, drug, and genetic data from current row
        processed_drug_name = process_drug_names(row['Drug(s)'])
        
        # Get phenotype and drug names for context
        phenotype = str(row['Phenotype(s)'])
        drug_name = str(row['Drug(s)'])
        
        # Get phenotype description from Wikipedia
        print(f"🔍 Fetching phenotype description from Wikipedia for row {index + 1}")
        phenotype_info = fetch_phenotype_description(phenotype)
        if phenotype_info['success']:
            phenotype_description = phenotype_info['description']
            print(f"  ✅ Phenotype description found: {phenotype_info['source']}")
        else:
            phenotype_description = f"Description not available for {phenotype}"
            print(f"  ❌ Phenotype description not found: {phenotype_info['description']}")
        
        # Initialize drug description - will be set from DrugBank or Wikipedia
        drug_description = f"Description not available for {drug_name}"
        
        # First try to get drug description from DrugBank
        print(f"🔬 Getting drug information from DrugBank for row {index + 1}")
        
        # Process drug name - take the first drug if multiple are listed
        drug_name_to_fetch = drug_name
        if ';' in drug_name:
            drug_name_to_fetch = drug_name.split(';')[0].strip()
            print(f"   Multiple drugs found, using first drug: {drug_name_to_fetch}")
        elif ',' in drug_name:
            drug_name_to_fetch = drug_name.split(',')[0].strip()
            print(f"   Multiple drugs found, using first drug: {drug_name_to_fetch}")
        
        drug_data = fetch_drug_data(drug_name_to_fetch)
        
        if drug_data and "error" not in drug_data and 'Summary' in drug_data and drug_data['Summary'] != "Summary not available":
            # Use DrugBank summary
            drug_description = drug_data['Summary']
            print(f"  ✅ Drug description from DrugBank: Summary section")
        else:
            # Fallback to Wikipedia if DrugBank doesn't have summary
            print(f"  ⚠️  No DrugBank summary available, trying Wikipedia...")
            drug_info = fetch_drug_description(drug_name)
            if drug_info['success']:
                drug_description = drug_info['description']
                print(f"  ✅ Drug description from Wikipedia: {drug_info['source']}")
            else:
                print(f"  ❌ Drug description not found in Wikipedia either")
        
        context = {
            'Phenotype': str(row['Phenotype(s)']),  # Phenotype data
            'Drug': processed_drug_name,  # Drug data with numbering and capitalization
            'Wikipedia_Info': phenotype_description,  # Phenotype description from Wikipedia
            'Wikipedia_Additional': drug_description,  # Drug description from Wikipedia
            'variant': str(row['Variant/Haplotypes']),  # Variant data
            'Gene': str(row['Gene']),  # Gene data
            'Genotype': str(row['Genotype/Allele']),  # Genotype data
            'Annotation_Text': str(row['Annotation Text']),  # Annotation data
            'Variant_Haplotypes': str(row['Variant/Haplotypes']),  # Table variant data
            'Genotype_Allele': str(row['Genotype/Allele']),  # Table genotype data
            'Phenotype_Category': str(row['Phenotype Category']),  # Table phenotype category
            'Level_of_Evidence': str(row['Level of Evidence']),  # Table level of evidence
            'Result': str(row['Result']) if 'Result' in row and pd.notna(row['Result']) else 'No result available'  # Result data
        }
        
        # Store the drug name for later use in detailed drug information
        row_drug_name = drug_name
        
        # Render template for this row
        try:
            doc = DocxTemplate(template_path)
            doc.render(context)
            
            # Apply color formatting to the genetic guidance box after rendering
            result_value = str(row['Result']) if 'Result' in row and pd.notna(row['Result']) else ''
            
            for paragraph in doc.docx.paragraphs:
                if 'Genetic Guidance for' in paragraph.text:
                    # Clear existing runs and recreate with proper formatting
                    paragraph.clear()
                    
                    # Add "Genetic Guidance for" part
                    guidance_run = paragraph.add_run('Genetic Guidance for ')
                    guidance_run.font.name = 'Times New Roman'
                    guidance_run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
                    guidance_run.font.size = Pt(11)
                    guidance_run.bold = True
                    
                    # Add drug name
                    processed_drug_name = process_drug_names(row['Drug(s)'])
                    drug_run = paragraph.add_run(processed_drug_name)
                    drug_run.font.name = 'Times New Roman'
                    drug_run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
                    drug_run.font.size = Pt(11)
                    drug_run.bold = True
                    
                    # Add newline
                    paragraph.add_run('\n')
                    
                    # Add result with color
                    result_run = paragraph.add_run(result_value)
                    result_run.font.name = 'Times New Roman'
                    result_run._element.rPr.rFonts.set(qn('w:ascii'), 'Times New Roman')
                    result_run.font.size = Pt(11)
                    result_run.bold = True
                    
                    # Set color based on result
                    if '✅ Administer normal dose' in result_value:
                        result_run.font.color.rgb = RGBColor(0, 128, 0)  # Green color
                    else:
                        result_run.font.color.rgb = RGBColor(255, 0, 0)  # Red color
            
            # Add detailed drug information from DrugBank (using data already fetched)
            try:
                print(f"🔬 Adding detailed drug information for row {index + 1}")
                
                                # Use the drug_data that was already fetched earlier
                if drug_data and "error" not in drug_data:
                    print(f"✅ Drug information received for row {index + 1}")
                    
                    # Find the insertion point (before "How to use the test results" section)
                    insertion_index = -1
                    for i, paragraph in enumerate(doc.docx.paragraphs):
                        if "How to use the test results" in paragraph.text:
                            insertion_index = i
                            break
                    
                    if insertion_index != -1:
                        ref_para = doc.docx.paragraphs[insertion_index]
                        
                        # Add detailed drug information title
                        title_para = ref_para.insert_paragraph_before(f"Detailed Drug Information: {processed_drug_name}")
                        title_run = title_para.runs[0]
                        title_run.font.name = "Times New Roman"
                        title_run.font.size = Pt(14)
                        title_run.bold = True
                        title_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                        # Add Summary section
                        if 'Summary' in drug_data and drug_data['Summary'] != "Summary not available":
                            summary_para = ref_para.insert_paragraph_before("Summary")
                            summary_run = summary_para.runs[0]
                            summary_run.font.name = "Times New Roman"
                            summary_run.font.size = Pt(12)
                            summary_run.bold = True
                            summary_run.font.color.rgb = RGBColor(0, 0, 0)
                            
                            summary_text_para = ref_para.insert_paragraph_before(drug_data['Summary'])
                            summary_text_run = summary_text_para.runs[0]
                            summary_text_run.font.name = "Times New Roman"
                            summary_text_run.font.size = Pt(11)
                            summary_text_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add Brand Names section
                        if 'Brand Names' in drug_data and drug_data['Brand Names'] != "Brand names not available":
                            brand_para = ref_para.insert_paragraph_before("Brand Names")
                            brand_run = brand_para.runs[0]
                            brand_run.font.name = "Times New Roman"
                            brand_run.font.size = Pt(12)
                            brand_run.bold = True
                            brand_run.font.color.rgb = RGBColor(0, 0, 0)
                            
                            brand_text_para = ref_para.insert_paragraph_before(drug_data['Brand Names'])
                            brand_text_run = brand_text_para.runs[0]
                            brand_text_run.font.name = "Times New Roman"
                            brand_text_run.font.size = Pt(11)
                            brand_text_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add Pharmacodynamics section
                        if 'Pharmacodynamics' in drug_data and drug_data['Pharmacodynamics'] != "Pharmacodynamics not available":
                            pharm_para = ref_para.insert_paragraph_before("Pharmacodynamics")
                            pharm_run = pharm_para.runs[0]
                            pharm_run.font.name = "Times New Roman"
                            pharm_run.font.size = Pt(12)
                            pharm_run.bold = True
                            pharm_run.font.color.rgb = RGBColor(0, 0, 0)
                            
                            pharm_text_para = ref_para.insert_paragraph_before(drug_data['Pharmacodynamics'])
                            pharm_text_run = pharm_text_para.runs[0]
                            pharm_text_run.font.name = "Times New Roman"
                            pharm_text_run.font.size = Pt(11)
                            pharm_text_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add ADME Properties
                        adme_fields = [
                            "Absorption", "Volume of distribution", "Protein binding",
                            "Metabolism", "Route of elimination", "Half-life", "Clearance", "Toxicity"
                        ]
                        
                        for field in adme_fields:
                            if field in drug_data and drug_data[field] != f"{field} not available":
                                field_para = ref_para.insert_paragraph_before(field)
                                field_run = field_para.runs[0]
                                field_run.font.name = "Times New Roman"
                                field_run.font.size = Pt(12)
                                field_run.bold = True
                                field_run.font.color.rgb = RGBColor(0, 0, 0)
                                
                                field_text_para = ref_para.insert_paragraph_before(drug_data[field])
                                field_text_run = field_text_para.runs[0]
                                field_text_run.font.name = "Times New Roman"
                                field_text_run.font.size = Pt(11)
                                field_text_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        
                        # Add General References section
                        if 'General References' in drug_data and isinstance(drug_data['General References'], list):
                            ref_para.insert_paragraph_before("")  # Add spacing
                            ref_title_para = ref_para.insert_paragraph_before("General References")
                            ref_title_run = ref_title_para.runs[0]
                            ref_title_run.font.name = "Times New Roman"
                            ref_title_run.font.size = Pt(12)
                            ref_title_run.bold = True
                            ref_title_run.font.color.rgb = RGBColor(0, 0, 0)
                            
                            for i, ref in enumerate(drug_data['General References'], 1):  # Show all references
                                ref_text = f"{i}. {ref['text']}"
                                if ref.get('link'):
                                    ref_text += f" - {ref['link']}"
                                
                                ref_item_para = ref_para.insert_paragraph_before(ref_text)
                                ref_item_run = ref_item_para.runs[0]
                                ref_item_run.font.name = "Times New Roman"
                                ref_item_run.font.size = Pt(10)
                                ref_item_para.paragraph_format.left_indent = Pt(20)
                        
                        print(f"✅ Detailed drug information added to row {index + 1}")
                    else:
                        print(f"❌ Could not find insertion point for detailed drug information in row {index + 1}")
                
                else:
                    print(f"❌ No valid drug data received for row {index + 1}")
                    
            except Exception as e:
                print(f"❌ Error adding detailed drug information for row {index + 1}: {e}")
            
            # Add page break after each phenotype (except the last one)
            if index < len(df) - 1:
                # Add a page break paragraph
                page_break_para = doc.docx.add_paragraph()
                page_break_para.add_run().add_break(WD_BREAK.PAGE)
            
            # Save to temporary file
            temp_file = f"temp_row_{index}.docx"
            doc.save(temp_file)
            temp_files.append(temp_file)
                    
        except Exception as e:
            print(f"❌ Error processing row {index + 1}: {e}")
            continue
    
    # Combine all temporary files
    try:
        from docxcompose.composer import Composer
        
        # Start with the first document
        combined_doc = Document(temp_files[0])
        composer = Composer(combined_doc)
        
        # Add the rest of the documents
        for temp_file in temp_files[1:]:
            doc = Document(temp_file)
            composer.append(doc)
        
        # Save the combined document
        composer.save(output_path)
        print(f"✅ Combined report generated: {output_path}")
        print(f"   Total rows processed: {len(df)}")
        
        # Clean up temporary files
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.remove(temp_file)
                
    except Exception as e:
        print(f"❌ Error saving combined report: {e}")
        # Clean up temporary files on error
        for temp_file in temp_files:
            if os.path.exists(temp_file):
                os.remove(temp_file)

def main():
    print("📊 Excel Report Generator")
    print("=" * 50)
    
    # Check if template exists and provide options
    if not os.path.exists(template_path):
        print("❌ Template not found!")
        print("Options:")
        print("1. Create new template (will overwrite any existing template)")
        print("2. Use backup template if available")
        
        choice = input("Enter choice (1 or 2): ").strip()
        
        if choice == "1":
            create_template()
            print(f"✅ Template created: {template_path}")
        elif choice == "2":
            backup_path = "genetic_report_template_backup.docx"
            if restore_template_from_backup(backup_path):
                print("✅ Using backup template")
            else:
                print("❌ No backup template found. Creating new template...")
                create_template()
        else:
            print("❌ Invalid choice. Creating new template...")
            create_template()
    else:
        print(f"✅ Template found: {template_path}")
        print("   Your manual changes will be preserved!")
    
    # Read Excel data
    df = read_excel_data()
    
    # Generate report
    generate_report(df)
    
    print("\n🎉 Report generation completed!")
    print("\n🎉 All processes completed!")

if __name__ == "__main__":
    main()
