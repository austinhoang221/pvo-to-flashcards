import pandas as pd
import re
def clean_text(text):
    """Remove HTML tags and convert to plain text"""
    if pd.isna(text):
        return ""
    # Remove HTML tags
    text = text.replace('<b>', '*').replace('</b>', '* ')
    text = re.sub(r'<[^>]+>', '', str(text))
    # Replace HTML entities
    text = text.replace('&nbsp;', ' ').replace('&amp;', '&')
    # Collapse multiple spaces and trim
    text = ' '.join(text.split())
    return text.strip()

def process_flashcards(input_file, output_file):
    # Load all sheets
    concepts_df = pd.read_excel(input_file, sheet_name='Concepts')
    examples_df = pd.read_excel(input_file, sheet_name='Examples')
    examples_links_df = pd.read_excel(input_file, sheet_name='Example Links')
    example_concepts_df = pd.read_excel(input_file, sheet_name='Example concepts')
    
    # Load metadata sheets if they exist
    tones_df = pd.read_excel(input_file, sheet_name='Tones') if 'Tones' in pd.ExcelFile(input_file).sheet_names else None
    modes_df = pd.read_excel(input_file, sheet_name='Modes') if 'Modes' in pd.ExcelFile(input_file).sheet_names else None
    dialects_df = pd.read_excel(input_file, sheet_name='Dialects') if 'Dialects' in pd.ExcelFile(input_file).sheet_names else None
    registers_df = pd.read_excel(input_file, sheet_name='Registers') if 'Registers' in pd.ExcelFile(input_file).sheet_names else None
    nuances_df = pd.read_excel(input_file, sheet_name='Nuances') if 'Nuances' in pd.ExcelFile(input_file).sheet_names else None

    # Create metadata mapping dictionaries
    def create_metadata_map(df, id_col='id', name_col='title'):
        return df.set_index(id_col)[name_col].to_dict() if df is not None else {}

    tone_map = create_metadata_map(tones_df)
    mode_map = create_metadata_map(modes_df)
    dialect_map = create_metadata_map(dialects_df)
    register_map = create_metadata_map(registers_df)
    nuance_map = create_metadata_map(nuances_df)

    # Create flashcards in the desired text format
    flashcard_entries = []
    
    for _, example_concept in example_concepts_df.iterrows():
        concept_id = example_concept['concept_id']
        example_id = example_concept['example_id']
        example_link_id = example_concept['example_link_id']
        
        # Get concept details
        concept = concepts_df[concepts_df['id'] == concept_id]
        if concept.empty:
            continue
        # Get example link details
        link = examples_links_df[examples_links_df['id'] == example_link_id]
        if link.empty:
            continue
        concept_title = concept.iloc[0]['title']
        concept_desc = concept.iloc[0]['description'] if 'description' in concept.columns else ''
        concept_link = link.iloc[0]['title'] if 'title' in concept.columns else ''
        
        # Get example details
        example = examples_df[examples_df['id'] == example_id]
        if example.empty:
            continue
            
        example_html = clean_text(example.iloc[0]['detail_html'])
        example_note = example.iloc[0]['note'] if 'note' in example.columns and pd.notna(example.iloc[0]['note']) else ''
        # Get metadata names (default to 'Neutral' if not specified)
        tone_name = tone_map.get(example.iloc[0]['tone_id'], 'Neutral') if 'tone_id' in example.columns else 'Neutral'
        mode_name = mode_map.get(example.iloc[0]['mode_id'], 'Neutral') if 'mode_id' in example.columns else 'Neutral'
        dialect_name = dialect_map.get(example.iloc[0]['dialect_id'], 'Neutral') if 'dialect_id' in example.columns else 'Neutral'
        register_name = register_map.get(example.iloc[0]['register_id'], 'Neutral') if 'register_id' in example.columns else 'Neutral'
        nuance_name = nuance_map.get(example.iloc[0]['nuance_id'], 'Neutral') if 'nuance_id' in example.columns else 'Neutral'
        
        # Build the flashcard text (Front;Back format)
        front = f"{concept_title}: {concept_desc}" if concept_desc else concept_title
        back = example_html.replace("&nbsp;", " ").replace("<br>", "\n").strip()
        
        if example_note:
            back += f"\n{example_note}"
        
        # Add metadata at the bottom
        metadata = f"Link: {concept_link}\nTone: {tone_name}\nMode: {mode_name}\nDialect: {dialect_name}\nRegister: {register_name}\nNuance: {nuance_name}"
        full_back = f"{back}\n{metadata}"
        
        # Combine into a single flashcard entry
        flashcard_entries.append(f"{front}/{full_back}")

    # Write to a text file with ";" separator
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(";\n".join(flashcard_entries) + ";")  # Ensures proper separation
    
    print(f"Flashcards saved to {output_file} (Text format)")

# Usage
input_excel = "pvo.xlsx"
output_txt = "quizlet_flashcards.txt"
process_flashcards(input_excel, output_txt)