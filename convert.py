import pandas as pd
from openpyxl import Workbook

def process_flashcards(input_file, output_file):
    # Load all sheets
    concepts_df = pd.read_excel(input_file, sheet_name='Concepts')
    examples_df = pd.read_excel(input_file, sheet_name='Examples')
    example_concepts_df = pd.read_excel(input_file, sheet_name='Example concepts')
    
    # Load metadata sheets if they exist (you may need to add these)
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
    neutral = "Neutral"
    # Create flashcards
    flashcards = []
    
    for _, example_concept in example_concepts_df.iterrows():
        concept_id = example_concept['concept_id']
        example_id = example_concept['example_id']
        
        # Get concept details
        concept = concepts_df[concepts_df['id'] == concept_id]
        if concept.empty:
            continue
            
        concept_title = concept.iloc[0]['title']
        concept_desc = concept.iloc[0]['description'] if 'description' in concept.columns else ''
        
        # Get example details
        example = examples_df[examples_df['id'] == example_id]
        if example.empty:
            continue
            
        example_html = example.iloc[0]['detail']
        example_note = example.iloc[0]['note'] if 'note' in example.columns and pd.notna(example.iloc[0]['note']) else ''
        
        # Get metadata names
        tone_name = tone_map.get(example.iloc[0]['tone_id'], neutral) if 'tone_id' in example.columns else neutral
        mode_name = mode_map.get(example.iloc[0]['mode_id'], neutral) if 'mode_id' in example.columns else neutral
        dialect_name = dialect_map.get(example.iloc[0]['dialect_id'], neutral) if 'dialect_id' in example.columns else neutral
        register_name = register_map.get(example.iloc[0]['register_id'], neutral) if 'register_id' in example.columns else neutral
        nuance_name = nuance_map.get(example.iloc[0]['nuance_id'], neutral) if 'nuance_id' in example.columns else neutral
        
        # Build front and back of card
        front = f"{concept_title}"
        if concept_desc:
            front += f": {concept_desc}"
            
        back = example_html.replace("&nbsp;", " ").strip()
        back += "\n"
        if example_note:
            back += f"{example_note + "\n" }"
        
        # Add metadata if available
        metadata_parts = []
        if tone_name: metadata_parts.append(f"Tone: {tone_name}")
        if mode_name: metadata_parts.append(f"Mode: {mode_name}")
        if dialect_name: metadata_parts.append(f"Dialect: {dialect_name}")
        if register_name: metadata_parts.append(f"Register: {register_name}")
        if nuance_name: metadata_parts.append(f"Nuance: {nuance_name}")
        
        if metadata_parts:
            back += "\n".join(metadata_parts) 
        
        flashcards.append({
            'Front': front,
            'Back': back + ";"
        })

    # Create output DataFrame
    output_df = pd.DataFrame(flashcards)
    
    # Save to Excel
    output_df.to_excel(output_file, index=False)
    print(f"Flashcards saved to {output_file}")

# Usage
input_excel = "pvo.xlsx"
output_excel = "quizlet_flashcards.xlsx"
process_flashcards(input_excel, output_excel)