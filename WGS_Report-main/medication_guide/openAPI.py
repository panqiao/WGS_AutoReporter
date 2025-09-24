import pandas as pd
from openai import OpenAI
import time

# Set up OpenAI API key
client = OpenAI(api_key="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")

def read_excel_and_process_all_rows():
    """Read the Excel file and process all rows"""
    try:
        # Read the Excel file
        excel_path = "/Users/durjoy/Documents/Ebovir/WGS_Report/leo-matched_snp_genotypes_results666_reordered.xlsx"
        df = pd.read_excel(excel_path)
        
        print(f"✅ Successfully read Excel file")
        print(f"📊 Total rows in file: {len(df)}")
        
        # Add Result column after Annotation Text column
        annotation_text_index = df.columns.get_loc('Annotation Text')
        df.insert(annotation_text_index + 1, 'Result', '')
        
        # Process each row
        results = []
        for index, row in df.iterrows():
            row_number = index + 1
            annotation_text = row['Annotation Text']
            
            # Skip if annotation text is empty or NaN
            if pd.isna(annotation_text) or annotation_text == '':
                print(f"⏭️ Skipping row {row_number} - no annotation text")
                continue
            
            print(f"\n📊 Processing row {row_number}/{len(df)}")
            print(f"   Annotation: {annotation_text[:100]}...")
            
            # Call OpenAI API
            api_result = call_openai_api(annotation_text, row_number)
            
            if api_result is None:
                print(f"❌ Could not get API response for row {row_number}")
                continue
            
            # Clean the result (remove number prefix)
            cleaned_result = clean_api_result(api_result)
            
            # Store result in DataFrame
            df.at[index, 'Result'] = cleaned_result
            
            # Store result for summary
            results.append({
                'row': row_number,
                'annotation': annotation_text,
                'ai_recommendation': cleaned_result
            })
            
            # Print individual result
            print(f"   AI Recommendation: {cleaned_result}")
            
            # Add a small delay to avoid rate limiting
            time.sleep(1)
        
        # Save the updated Excel file
        df.to_excel(excel_path, index=False)
        print(f"\n💾 Updated Excel file saved with results")
        
        return results
        
    except Exception as e:
        print(f"❌ Error reading Excel file: {e}")
        return None

def clean_api_result(api_result):
    """Clean the API result by removing number prefix"""
    # Remove common number patterns like "1.", "2.", etc.
    import re
    cleaned = re.sub(r'^\d+\.\s*', '', api_result.strip())
    return cleaned

def call_openai_api(annotation_text, row_number):
    """Call OpenAI API with the annotation text"""
    try:
        # Create the prompt
        prompt = f"""
        Please analyze the following genetic annotation text and select the most appropriate option from the following list, only return the option, nothing else:

        Annotation Text: {annotation_text}

        Please provide:
        1. 🚫 Consider alternative,
        2. ⚠️ Use with caution,
        3. ⬆️ Increase dose,
        4. ⬇️ Decrease dose,
        5. ✅ Administer normal dose,
        6. ❓ Not clinically actionable,
        """
        
        print(f"🤖 Calling OpenAI API for row {row_number}...")
        
        # Make the API call
        response = client.chat.completions.create(
            model="gpt-4",
            messages=[
                {"role": "system", "content": "You are a highly knowledgeable pharmacology assistant specialized in clinical pharmacokinetics, pharmacogenomics, and evidence-based therapeutics."},
                {"role": "user", "content": prompt}
            ],
            max_tokens=1000,
            temperature=0.3
        )
        
        # Extract the response
        api_response = response.choices[0].message.content
        
        print(f"✅ API call successful for row {row_number}!")
        return api_response
        
    except Exception as e:
        print(f"❌ Error calling OpenAI API for row {row_number}: {e}")
        return None

def main():
    print("🧬 Genetic Annotation Analysis with OpenAI")
    print("=" * 50)
    
    # Process all rows
    results = read_excel_and_process_all_rows()
    
    if results is None:
        print("❌ Could not process Excel file")
        return
    
    # Print summary
    print("\n" + "=" * 50)
    print("📋 SUMMARY OF ALL OPENAI API ANALYSIS RESULTS")
    print("=" * 50)
    
    for result in results:
        print(f"\nRow {result['row']}:")
        print(f"Recommendation: {result['ai_recommendation']}")
    
    print(f"\n✅ Processed {len(results)} rows successfully")
    print("=" * 50)

if __name__ == "__main__":
    main()
