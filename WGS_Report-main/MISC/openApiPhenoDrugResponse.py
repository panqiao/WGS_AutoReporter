import pandas as pd
import requests
import json
import os
from typing import Dict, Any

# File paths
excel_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/leo-matched_snp_genotypes_results666_reordered.xlsx"

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

def make_openapi_call(phenotype: str, drug_name: str) -> Dict[Any, Any]:
    """
    Make OpenAPI call with the phenotype and drug description prompt
    """
    
    # Process drug names for formatting
    processed_drug_name = process_drug_names(drug_name)
    
    # The prompt template for JSON-only response
    prompt = f"""
You are a highly knowledgeable pharmacology assistant specialized in clinical pharmacokinetics, pharmacogenomics, and evidence-based therapeutics. Please provide concise drug description and phenotype description for **{phenotype}**, **{processed_drug_name}** in the following JSON format only. Do not include any other text or explanations outside the JSON structure.

IMPORTANT: Keep each description to 4-5 sentences maximum. Be concise but informative.

```json
{{
  "content": [
    {{
      "type": "paragraph",
      "order": 1,
      "title": "1. Phenotype Description",
      "text": "Concise phenotype description (4-5 sentences) for {phenotype}"
    }},
    {{
      "type": "paragraph",
      "order": 2,
      "title": "2. Drug Description",
      "text": "Concise drug overview (4-5 sentences) including mechanism of action, approved indications, contraindications, and black box warnings for {processed_drug_name}"
    }}
  ]
}}
```

Provide ONLY the JSON response with concise, evidence-based information (4-5 sentences each) for {phenotype} and {processed_drug_name}.
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
                "content": "You are a highly knowledgeable pharmacology assistant specialized in clinical pharmacokinetics, pharmacogenomics, and evidence-based therapeutics. Provide comprehensive, accurate, and up-to-date drug and phenotype information following the specified format exactly."
            },
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 3000,
        "temperature": 0.2
    }
    
    try:
        print(f"🔍 Making API call for phenotype: {phenotype}, drugs: {processed_drug_name}")
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

def print_response_data(response_data: Dict[Any, Any], phenotype: str, drug_name: str):
    """Print the API response data to console"""
    try:
        print(f"\n📋 Response for Phenotype: {phenotype}")
        print(f"📋 Response for Drugs: {drug_name}")
        print("=" * 80)
        
        # Extract and print the content
        content = extract_text_content(response_data)
        if content and content != "No content found in response":
            print("JSON Response:")
            print(content)
            print("=" * 80)
            return content
        else:
            print("❌ No valid content in response")
            return None
    except Exception as e:
        print(f"❌ Error printing response: {e}")
        return None

def process_all_rows():
    """Process all rows and return the data"""
    print("🔬 Phenotype and Drug Information API Caller")
    print("=" * 60)
    
    # Read Excel data
    df = read_excel_data()
    if df is None:
        return None
    
    all_responses = []
    
    # Process each row
    for index, row in df.iterrows():
        phenotype = str(row['Phenotype(s)'])
        drug_name = str(row['Drug(s)'])
        
        print(f"\n📄 Processing row {index + 1}/{len(df)}")
        print(f"   Phenotype: {phenotype}")
        print(f"   Drugs: {drug_name}")
        
        # Make API call
        response = make_openapi_call(phenotype, drug_name)
        
        if response:
            # Extract and print content
            content = extract_text_content(response)
            if content and content != "No content found in response":
                print(f"✅ API Response Content received")
                
                # Print the response data
                response_data = print_response_data(response, phenotype, drug_name)
                
                # Store the response data
                all_responses.append({
                    'row_index': index,
                    'phenotype': phenotype,
                    'drug_name': drug_name,
                    'response_data': response_data,
                    'full_response': response
                })
            else:
                print(f"❌ No valid content received")
        else:
            print(f"❌ No response received from API")
    
    print(f"\n🎉 Completed processing all 5 rows!")
    return all_responses

def main():
    """Main function that processes all rows"""
    return process_all_rows()

if __name__ == "__main__":
    main()
