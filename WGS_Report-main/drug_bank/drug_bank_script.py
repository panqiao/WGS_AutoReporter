from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import time
import json
import re
import pandas as pd
import sys
import os

def find_drugbank_id(drug_name, tsv_file="drugbank.tsv"):
    """
    Find the DrugBank ID for a given drug name from the TSV file
    """
    try:
        # Read the TSV file
        df = pd.read_csv(tsv_file, sep='\t')
        
        # Search for the drug name (case-insensitive)
        drug_name_lower = drug_name.lower().strip()
        
        # Try exact match first
        exact_match = df[df['name'].str.lower() == drug_name_lower]
        if not exact_match.empty:
            return exact_match.iloc[0]['drugbank_id']
        
        # Try partial match
        partial_match = df[df['name'].str.lower().str.contains(drug_name_lower, na=False)]
        if not partial_match.empty:
            print(f"Found partial match: {partial_match.iloc[0]['name']}")
            return partial_match.iloc[0]['drugbank_id']
        
        # If no match found, return None
        print(f"Drug '{drug_name}' not found in DrugBank database")
        return None
        
    except Exception as e:
        print(f"Error reading TSV file: {e}")
        return None

def extract_references_with_tags(element):
    """
    Extract reference numbers and format them with <ref> tags
    """
    if not element:
        return ""
    
    try:
        # Get the text content first
        text = element.text.strip()
        
        # Find all reference elements
        ref_elements = element.find_elements(By.XPATH, ".//sup[@class='text-reference-group']")
        
        if not ref_elements:
            return text
        
        # Process each reference element to replace numbers with <ref> tags
        for ref_element in ref_elements:
            try:
                # Get the reference number from the sup element
                ref_text = ref_element.text.strip()
                if ref_text:
                    # Handle comma-separated references (e.g., "16,20,21")
                    if ',' in ref_text:
                        # Split by comma and format each reference
                        ref_numbers = [num.strip() for num in ref_text.split(',')]
                        formatted_refs = ','.join([f"<ref>{num}</ref>" for num in ref_numbers])
                        # Replace the original text with formatted references
                        text = text.replace(ref_text, formatted_refs)
                    else:
                        # Single reference number
                        text = text.replace(ref_text, f"<ref>{ref_text}</ref>")
            except:
                continue
        
        # Clean up any remaining issues: fix nested ref tags and unformatted references
        import re
        
        # Remove nested ref tags (e.g., <ref><ref>16</ref></ref> -> <ref>16</ref>)
        text = re.sub(r'<ref><ref>(\d+)</ref></ref>', r'<ref>\1</ref>', text)
        
        # Fix any remaining unformatted references (e.g., ",20" -> ",<ref>20</ref>")
        text = re.sub(r',(\d+)(?=\s|$|\.)', r',<ref>\1</ref>', text)
        
        # Fix references at the end of sentences
        text = re.sub(r'(\d+)(?=\s*$)', r'<ref>\1</ref>', text)
        
        return text
        
    except Exception as e:
        # Fallback to regular text extraction
        return element.text.strip()

def extract_general_references(element):
    """
    Extract General References and format them as a structured list
    """
    if not element:
        return ""
    
    try:
        # Find the general references section
        general_refs_dt = element.find_element(By.XPATH, "//dt[@id='general-references']")
        general_refs_dd = general_refs_dt.find_element(By.XPATH, "following-sibling::dd[1]")
        
        # Find all list items in the references
        ref_items = general_refs_dd.find_elements(By.TAG_NAME, "li")
        
        general_references = []
        for item in ref_items:
            try:
                # Get the reference ID
                ref_id = item.get_attribute("id")
                
                # Get the text content (excluding the [Article] or [Link] part)
                ref_text = item.text
                
                # Clean up the text by removing the [Article], [Link], or [File] part
                ref_text = re.sub(r'\s*\[(Article|Link|File)\]\s*$', '', ref_text)
                
                # Get the link if available
                try:
                    link_element = item.find_element(By.TAG_NAME, "a")
                    link_url = link_element.get_attribute("href")
                except:
                    link_url = ""
                
                general_references.append({
                    "id": ref_id,
                    "text": ref_text.strip(),
                    "link": link_url
                })
            except Exception as e:
                # If we can't parse an item, just get the text
                try:
                    ref_text = item.text
                    ref_text = re.sub(r'\s*\[(Article|Link|File)\]\s*$', '', ref_text)
                    general_references.append({
                        "id": f"ref_{len(general_references) + 1}",
                        "text": ref_text.strip(),
                        "link": ""
                    })
                except:
                    continue
        
        # Format references as a structured string
        if general_references:
            ref_strings = []
            for ref in general_references:
                ref_str = f"{ref['id']}: {ref['text']}"
                if ref['link']:
                    ref_str += f" | {ref['link']}"
                ref_strings.append(ref_str)
            
            return " || ".join(ref_strings)
        else:
            return "No general references available"
            
    except NoSuchElementException:
        return "General references not available"
    except Exception as e:
        return f"Error extracting references: {str(e)}"

def fetch_drug_data(drug_name, tsv_file="drugbank.tsv"):
    """
    Fetch drug data from DrugBank for a given drug name
    """
    # Find the DrugBank ID for the drug
    drugbank_id = find_drugbank_id(drug_name, tsv_file)
    
    if not drugbank_id:
        return {"error": f"Drug '{drug_name}' not found in DrugBank database"}
    
    print(f"Found DrugBank ID: {drugbank_id} for drug: {drug_name}")
    
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

    driver = webdriver.Chrome(options=options)

    try:
        url = f"https://go.drugbank.com/drugs/{drugbank_id}"
        print(f"Accessing URL: {url}")
        driver.get(url)

        # Wait for potential Cloudflare protection
        print("Waiting for page to load...")
        time.sleep(3)
        if "Just a moment" in driver.title:
            print("Still on Cloudflare protection page. Waiting longer...")
            time.sleep(30)

        print(f"Page title: {driver.title}")
        data = {}

        # --- Identification ---
        try:
            drugbank_id_element = driver.find_element(By.XPATH, "//dt[contains(.,'DrugBank ID')]/following-sibling::dd[1]").text
        except NoSuchElementException:
            drugbank_id_element = drugbank_id

        try:
            generic_name = driver.find_element(By.XPATH, "//dt[contains(.,'Generic Name')]/following-sibling::dd[1]").text
        except NoSuchElementException:
            generic_name = drug_name

        data['Identification'] = {
            'DrugBank ID': drugbank_id_element,
            'Generic Name': generic_name
        }

        # --- Summary ---
        try:
            summary_dt = driver.find_element(By.XPATH, "//dt[@id='summary']")
            summary_dd = summary_dt.find_element(By.XPATH, "following-sibling::dd[1]")
            data['Summary'] = summary_dd.text.strip()
        except NoSuchElementException:
            try:
                # Fallback: Try overview section
                summary = driver.find_element(By.XPATH, "//section[@id='overview']//div[contains(@class,'description')]").text
                data['Summary'] = summary
            except NoSuchElementException:
                data['Summary'] = "Summary not available"

        # --- Indication ---
        try:
            indication_dt = driver.find_element(By.XPATH, "//dt[@id='indication']")
            indication_dd = indication_dt.find_element(By.XPATH, "following-sibling::dd[1]")
            
            # Extract text with reference numbers formatted as <ref> tags
            indication_text = extract_references_with_tags(indication_dd)
            
            # Remove the promotional content completely
            promotional_patterns = [
                "Build, train, & validate predictive machine-learning models with structured datasets. SEE HOW",
                "Build, train, & validate predictive machine-learning models with structured datasets.",
                "SEE HOW"
            ]
            
            for pattern in promotional_patterns:
                if pattern in indication_text:
                    indication_text = indication_text.replace(pattern, "").strip()
                    print(f"Removed promotional content: {pattern[:50]}...")
            
            # Additional cleanup: remove any remaining lines that contain promotional keywords
            lines = indication_text.split('\n')
            clean_lines = []
            for line in lines:
                line = line.strip()
                if not line:
                    continue
                
                # Skip lines that contain promotional keywords
                if any(keyword in line.lower() for keyword in ['build, train', 'machine-learning', 'structured datasets', 'see how']):
                    continue
                
                clean_lines.append(line)
            
            # Join clean lines back together
            clean_indication = ' '.join(clean_lines).strip()
            data['Indication'] = clean_indication
            
        except NoSuchElementException:
            data['Indication'] = "Indication not available"

        # --- Mechanism of Action ---
        try:
            mechanism_dt = driver.find_element(By.XPATH, "//dt[@id='mechanism-of-action']")
            mechanism_dd = mechanism_dt.find_element(By.XPATH, "following-sibling::dd[1]")
            
            # Remove the HTML table first, then get clean text
            try:
                # Find and remove the target table if it exists
                target_table = mechanism_dd.find_element(By.XPATH, ".//table[@id='drug-moa-target-table']")
                if target_table:
                    # Remove the table from the DOM temporarily to get clean text
                    driver.execute_script("arguments[0].remove();", target_table)
                    print("Removed target table from Mechanism of Action")
            except NoSuchElementException:
                # No table found, continue normally
                pass
            
            # Extract text with reference numbers formatted as <ref> tags
            data['Mechanism of Action'] = extract_references_with_tags(mechanism_dd)
            
        except NoSuchElementException:
            data['Mechanism of Action'] = "Mechanism of action not available"

        # --- Brand Names ---
        try:
            brand_names = driver.find_element(By.XPATH, "//dt[contains(.,'Brand Names')]/following-sibling::dd[1]").text
            data['Brand Names'] = brand_names
        except NoSuchElementException:
            data['Brand Names'] = "Brand names not available"

        # --- Pharmacology & ADME ---
        try:
            # Pharmacodynamics - using the correct selector based on page structure
            try:
                pharmacodynamics_dt = driver.find_element(By.XPATH, "//dt[@id='pharmacodynamics']")
                pharmacodynamics_dd = pharmacodynamics_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                data['Pharmacodynamics'] = extract_references_with_tags(pharmacodynamics_dd)
            except NoSuchElementException:
                data['Pharmacodynamics'] = "Pharmacodynamics not available"

            # ADME Properties - using individual dt/dd elements
            properties = {}
            
            # Absorption
            try:
                absorption_dt = driver.find_element(By.XPATH, "//dt[@id='absorption']")
                absorption_dd = absorption_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Absorption'] = extract_references_with_tags(absorption_dd)
            except NoSuchElementException:
                properties['Absorption'] = "Absorption not available"

            # Volume of distribution
            try:
                vd_dt = driver.find_element(By.XPATH, "//dt[@id='volume-of-distribution']")
                vd_dd = vd_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Volume of distribution'] = extract_references_with_tags(vd_dd)
            except NoSuchElementException:
                properties['Volume of distribution'] = "Volume of distribution not available"

            # Protein binding
            try:
                protein_dt = driver.find_element(By.XPATH, "//dt[@id='protein-binding']")
                protein_dd = protein_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Protein binding'] = extract_references_with_tags(protein_dd)
            except NoSuchElementException:
                properties['Protein binding'] = "Protein binding not available"

            # Metabolism
            try:
                metabolism_dt = driver.find_element(By.XPATH, "//dt[@id='metabolism']")
                metabolism_dd = metabolism_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Metabolism'] = extract_references_with_tags(metabolism_dd)
            except NoSuchElementException:
                properties['Metabolism'] = "Metabolism not available"

            # Route of elimination
            try:
                elimination_dt = driver.find_element(By.XPATH, "//dt[@id='route-of-elimination']")
                elimination_dd = elimination_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Route of elimination'] = extract_references_with_tags(elimination_dd)
            except NoSuchElementException:
                properties['Route of elimination'] = "Route of elimination not available"

            # Half-life
            try:
                half_life_dt = driver.find_element(By.XPATH, "//dt[@id='half-life']")
                half_life_dd = half_life_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Half-life'] = extract_references_with_tags(half_life_dd)
            except NoSuchElementException:
                properties['Half-life'] = "Half-life not available"

            # Clearance
            try:
                clearance_dt = driver.find_element(By.XPATH, "//dt[@id='clearance']")
                clearance_dd = clearance_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Clearance'] = extract_references_with_tags(clearance_dd)
            except NoSuchElementException:
                properties['Clearance'] = "Clearance not available"

            # Toxicity
            try:
                toxicity_dt = driver.find_element(By.XPATH, "//dt[@id='toxicity']")
                toxicity_dd = toxicity_dt.find_element(By.XPATH, "following-sibling::dd[1]")
                properties['Toxicity'] = extract_references_with_tags(toxicity_dd)
            except NoSuchElementException:
                properties['Toxicity'] = "Toxicity not available"

            # Add all properties to data
            for field in [
                "Absorption", "Volume of distribution", "Protein binding",
                "Metabolism", "Route of elimination", "Half-life", "Clearance", "Toxicity"
            ]:
                data[field] = properties.get(field, f"{field} not available")

        except Exception as e:
            print(f"Pharmacology section error: {e}")
            data['Pharmacodynamics'] = "Pharmacodynamics not available"
            for field in [
                "Absorption", "Volume of distribution", "Protein binding",
                "Metabolism", "Route of elimination", "Half-life", "Clearance", "Toxicity"
            ]:
                data[field] = f"{field} not available"

        # --- General References ---
        try:
            data['General References'] = extract_general_references(driver)
        except Exception as e:
            data['General References'] = "General references not available"

        return data

    except Exception as e:
        print(f"Error occurred: {e}")
        return {"error": str(e)}

    finally:
        driver.quit()

def read_drugs_from_excel(excel_file="preprocessed_data_v1.xlsx", max_rows=None):
    """
    Read drug names from the Excel file (process all rows by default)
    """
    try:
        # Read the Excel file
        df = pd.read_excel(excel_file)
        
        # Check if 'Drug(s)' column exists
        if 'Drug(s)' not in df.columns:
            print(f"Error: 'Drug(s)' column not found in {excel_file}")
            print(f"Available columns: {df.columns.tolist()}")
            return []
        
        # Get drug names from all rows or specified number of rows
        if max_rows:
            drug_names = df['Drug(s)'].head(max_rows).dropna().astype(str).str.strip()
            print(f"Processing first {max_rows} rows from Excel file")
        else:
            drug_names = df['Drug(s)'].dropna().astype(str).str.strip()
            print(f"Processing ALL rows from Excel file")
        
        drug_names = drug_names[drug_names != '']
        
        # Process drug names: if multiple drugs separated by semicolon, take only the first one
        processed_drug_names = []
        for drug_name in drug_names:
            if ';' in drug_name:
                # Split by semicolon and take the first drug
                first_drug = drug_name.split(';')[0].strip()
                processed_drug_names.append(first_drug)
                print(f"Multiple drugs found: '{drug_name}' -> Using first drug: '{first_drug}'")
            else:
                processed_drug_names.append(drug_name)
        
        # Remove duplicates
        unique_drug_names = list(dict.fromkeys(processed_drug_names))
        
        print(f"Found {len(unique_drug_names)} unique drugs in {excel_file}")
        return unique_drug_names
        
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return []

if __name__ == "__main__":
    # Check if Excel file path is provided as command line argument
    if len(sys.argv) > 1:
        excel_file = sys.argv[1]
        print(f"Reading drugs from Excel file: {excel_file}")
    else:
        # Default to the preprocessed data file
        # Try to find the file in the current directory or parent directory
        possible_paths = [
            "preprocessed_data_v1.xlsx",
            "../preprocessed_data_v1.xlsx",
            os.path.join(os.path.dirname(__file__), "..", "preprocessed_data_v1.xlsx")
        ]
        
        excel_file = None
        for path in possible_paths:
            if os.path.exists(path):
                excel_file = path
                break
        
        if excel_file is None:
            print("Error: Could not find preprocessed_data_v1.xlsx")
            print("Please provide the full path to the Excel file as command line argument")
            print("Or place the Excel file in the same directory as this script")
            sys.exit(1)
        
        print(f"Using Excel file: {excel_file}")
    
    # Check if file exists
    if not os.path.exists(excel_file):
        print(f"Error: Excel file '{excel_file}' not found")
        print("Please provide a valid Excel file path as command line argument")
        sys.exit(1)
    
    # Read drugs from Excel file
    drug_names = read_drugs_from_excel(excel_file)
    
    if not drug_names:
        print("No drugs found in Excel file. Exiting.")
        sys.exit(1)
    
    print(f"\nProcessing all rows from Excel file for complete dataset...")
    
    # Process each drug
    all_results = []
    for i, drug_name in enumerate(drug_names):
        print(f"\n{'='*50}")
        print(f"Processing drug {i+1}/{len(drug_names)}: {drug_name}")
        print(f"{'='*50}")
        
        try:
            result = fetch_drug_data(drug_name)
            
            # Flatten the result data for CSV
            flat_result = {
                'Drug_Name': drug_name
            }
            
            # Extract data from nested structure
            if 'Identification' in result:
                flat_result['DrugBank_ID'] = result['Identification'].get('DrugBank ID', '')
            
            if 'Summary' in result:
                flat_result['Summary'] = result['Summary']
            
            if 'Indication' in result:
                flat_result['Indication'] = result['Indication']
            
            if 'Mechanism of Action' in result:
                flat_result['Mechanism_of_Action'] = result['Mechanism of Action']
            
            if 'Pharmacodynamics' in result:
                flat_result['Pharmacodynamics'] = result['Pharmacodynamics']
            
            if 'Absorption' in result:
                flat_result['Absorption'] = result['Absorption']
            
            if 'Volume of distribution' in result:
                flat_result['Volume_of_Distribution'] = result['Volume of distribution']
            
            if 'Protein binding' in result:
                flat_result['Protein_Binding'] = result['Protein binding']
            
            if 'Metabolism' in result:
                flat_result['Metabolism'] = result['Metabolism']
            
            if 'Route of elimination' in result:
                flat_result['Route_of_Elimination'] = result['Route of elimination']
            
            if 'Half-life' in result:
                flat_result['Half_Life'] = result['Half-life']
            
            if 'Clearance' in result:
                flat_result['Clearance'] = result['Clearance']
            
            if 'Toxicity' in result:
                flat_result['Toxicity'] = result['Toxicity']
            
            if 'General References' in result:
                flat_result['General_References'] = result['General References']
            
            all_results.append(flat_result)
            
            # Add a small delay between requests to be respectful to the server
            if i < len(drug_names) - 1:  # Don't delay after the last drug
                time.sleep(2)
                
        except Exception as e:
            print(f"Error processing drug '{drug_name}': {e}")
            # Add error record to CSV without Status column
            error_result = {
                'Drug_Name': drug_name,
                'Error_Message': str(e)
            }
            all_results.append(error_result)
    
    # Save all results to a CSV file
    output_file = "drug_bank_results.csv"
    
    if all_results:
        # Create DataFrame and save to CSV
        df = pd.DataFrame(all_results)
        df.to_csv(output_file, index=False, encoding='utf-8')
        print(f"\nResults saved to CSV file: {output_file}")
        
        # Display first few rows
        print("\nFirst few rows of the CSV:")
        print(df.head())
    
    print(f"\n{'='*50}")
    print(f"Processing complete! Results saved to {output_file}")
    print(f"Processed {len(drug_names)} drugs from all rows")
    print(f"{'='*50}")
    
    # Print summary
    successful = sum(1 for result in all_results if 'DrugBank_ID' in result and result.get('DrugBank_ID'))
    failed = len(drug_names) - successful
    print(f"Successful: {successful}")
    print(f"Failed: {failed}")
