import wikipedia
import re

def search_wikipedia_page(search_term):
    """Smart search for Wikipedia pages, handling disambiguation"""
    try:
        # Try direct search first
        return wikipedia.page(search_term)
    except wikipedia.exceptions.DisambiguationError as e:
        # If there's disambiguation, look for medical/health-related options
        medical_keywords = ['disease', 'disorder', 'syndrome', 'condition', 'medicine', 'drug', 'medication', 'pharmaceutical', 'health', 'medical', 'psychiatry', 'psychology', 'mental', 'clinical']
        
        # First, try to find exact medical matches
        for option in e.options:
            option_lower = option.lower()
            if any(keyword in option_lower for keyword in medical_keywords):
                try:
                    page = wikipedia.page(option)
                    # Double-check the page content to ensure it's medical
                    if any(medical_keyword in page.summary.lower() for medical_keyword in medical_keywords):
                        return page
                except:
                    continue
        
        # If no medical option found, try to find pages that contain medical content
        for option in e.options:
            try:
                page = wikipedia.page(option)
                # Check if the page content is medical-related
                if any(medical_keyword in page.summary.lower() for medical_keyword in medical_keywords):
                    return page
            except:
                continue
        
        # If still no medical option, use the first option but log a warning
        print(f"⚠️  Warning: No clear medical page found for '{search_term}'. Using first option: {e.options[0]}")
        return wikipedia.page(e.options[0])
    except wikipedia.exceptions.PageError:
        return None

def fetch_phenotype_description(phenotype_name):
    """
    Fetch phenotype description from Wikipedia
    
    Args:
        phenotype_name (str): Name of the phenotype/disease/disorder
        
    Returns:
        dict: Dictionary containing description and metadata
    """
    try:
        # Set language to English
        wikipedia.set_lang('en')
        
        print(f"🔍 Searching Wikipedia for phenotype: {phenotype_name}")
        
        # Handle common medical terms that might have disambiguation issues
        search_variations = [phenotype_name]
        
        # Add common medical variations for better search
        if phenotype_name.lower() == "depression":
            search_variations = ["Major depressive disorder", "Clinical depression", "Depression (mood)", "Depression"]
        elif phenotype_name.lower() == "anxiety":
            search_variations = ["Anxiety disorder", "Anxiety", "Generalized anxiety disorder"]
        elif phenotype_name.lower() == "bipolar":
            search_variations = ["Bipolar disorder", "Bipolar", "Manic depression"]
        elif phenotype_name.lower() == "adhd":
            search_variations = ["Attention deficit hyperactivity disorder", "ADHD", "Attention deficit disorder"]
        
        # Try each search variation
        page = None
        for search_term in search_variations:
            try:
                page = search_wikipedia_page(search_term)
                if page:
                    print(f"  → Found page using search term: '{search_term}'")
                    break
            except:
                continue
        
        if not page:
            # Fallback to original search
            page = search_wikipedia_page(phenotype_name)
        if not page:
            return {
                'success': False,
                'description': f"Could not find Wikipedia page for '{phenotype_name}'",
                'source': 'Error',
                'page_title': 'N/A'
            }
        
        print(f"  → Found page: {page.title}")
        
        # Try to get specific sections in order of preference
        sections_to_try = [
            'Signs and symptoms',
            'Symptoms',
            'Clinical features',
            'Overview',
            'Description',
            'Introduction'
        ]
        
        content = None
        source_section = None
        
        for section in sections_to_try:
            try:
                content = page.section(section)
                if content and len(content.strip()) > 50:  # Ensure meaningful content
                    source_section = section
                    break
            except:
                continue
        
        # If no specific section found, use summary
        if not content:
            content = page.summary
            source_section = 'Summary'
        
        # Clean up the content
        if content:
            # Remove excessive whitespace and newlines
            content = re.sub(r'\n+', ' ', content)
            content = re.sub(r'\s+', ' ', content)
            content = content.strip()
            
            # Limit to 4-5 sentences
            sentences = re.split(r'[.!?]+', content)
            # Filter out empty sentences and take first 4-5 meaningful sentences
            meaningful_sentences = [s.strip() for s in sentences if s.strip() and len(s.strip()) > 10]
            if len(meaningful_sentences) > 5:
                content = '. '.join(meaningful_sentences[:5]) + '.'
            else:
                content = '. '.join(meaningful_sentences) + '.'
        
        return {
            'success': True,
            'description': content,
            'source': f'Section: {source_section}',
            'page_title': page.title
        }
        
    except Exception as e:
        return {
            'success': False,
            'description': f"Error fetching phenotype description: {str(e)}",
            'source': 'Error',
            'page_title': 'N/A'
        }

def fetch_drug_description(drug_name):
    """
    Fetch drug description from Wikipedia
    
    Args:
        drug_name (str): Name of the drug/medication
        
    Returns:
        dict: Dictionary containing description and metadata
    """
    try:
        # Set language to English
        wikipedia.set_lang('en')
        
        print(f"💊 Searching Wikipedia for drug: {drug_name}")
        
        # Search for the drug page
        page = search_wikipedia_page(drug_name)
        if not page:
            return {
                'success': False,
                'description': f"Could not find Wikipedia page for '{drug_name}'",
                'source': 'Error',
                'page_title': 'N/A'
            }
        
        print(f"  → Found page: {page.title}")
        
        # Try to get specific sections in order of preference
        sections_to_try = [
            'Medical uses',
            'Uses',
            'Indications',
            'Pharmacology',
            'Mechanism of action',
            'Overview',
            'Description',
            'Introduction'
        ]
        
        content = None
        source_section = None
        
        for section in sections_to_try:
            try:
                content = page.section(section)
                if content and len(content.strip()) > 50:  # Ensure meaningful content
                    source_section = section
                    break
            except:
                continue
        
        # If no specific section found, use summary
        if not content:
            content = page.summary
            source_section = 'Summary'
        
        # Clean up the content
        if content:
            # Remove excessive whitespace and newlines
            content = re.sub(r'\n+', ' ', content)
            content = re.sub(r'\s+', ' ', content)
            content = content.strip()
            
            # Limit to 4-5 sentences
            sentences = re.split(r'[.!?]+', content)
            # Filter out empty sentences and take first 4-5 meaningful sentences
            meaningful_sentences = [s.strip() for s in sentences if s.strip() and len(s.strip()) > 10]
            if len(meaningful_sentences) > 5:
                content = '. '.join(meaningful_sentences[:5]) + '.'
            else:
                content = '. '.join(meaningful_sentences) + '.'
        
        return {
            'success': True,
            'description': content,
            'source': f'Section: {source_section}',
            'page_title': page.title
        }
        
    except Exception as e:
        return {
            'success': False,
            'description': f"Error fetching drug description: {str(e)}",
            'source': 'Error',
            'page_title': 'N/A'
        }

# Test functions if run directly
if __name__ == "__main__":
    print("Wikipedia Data Fetcher - Test Mode")
    print("=" * 50)
    
    # Test phenotype function
    print("\n🧬 Testing Phenotype Description:")
    phenotype_result = fetch_phenotype_description("Schizophrenia")
    print(f"Success: {phenotype_result['success']}")
    print(f"Source: {phenotype_result['source']}")
    print(f"Description: {phenotype_result['description'][:200]}...")
    
    # Test drug function
    print("\n💊 Testing Drug Description:")
    drug_result = fetch_drug_description("Risperidone")
    print(f"Success: {drug_result['success']}")
    print(f"Source: {drug_result['source']}")
    print(f"Description: {drug_result['description'][:200]}...")
