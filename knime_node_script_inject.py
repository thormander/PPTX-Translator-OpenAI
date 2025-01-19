import knime.scripting.io as knio
import re
from pptx import Presentation

# Put dataframe into hashmap for easy retrieval of text that corresponds to id
df = knio.input_tables[0].to_pandas()
translation_hashmap = df.set_index('id')['Response'].to_dict()

####################################### Parsing Logic ##########################################
# Slight modification from first python script, here we just want to fetch from our map and insert
# the translated text


input_file = knio.flow_variables["tempPPTXpath"]

def contains_meaningful_content(text):
    return bool(re.search(r'[a-zA-Z0-9]', text))

def extract_shape_text(shape) -> None:
    """Extract text from shape and add a unique ID."""
    if hasattr(shape, "text_frame") and shape.text_frame:
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                if run.text.strip() and contains_meaningful_content(run.text):
                    extracted_id = run.text.strip() # get id

                    try:           
                        extracted_id = int(extracted_id) # convert
                        run.text = translation_hashmap.get(extracted_id,"")
                    except:
                        run.text = ""

def extract_table_text(table) -> None:
    """Extract text from table cells and add a unique ID."""
    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.text_frame.paragraphs:
                for run in paragraph.runs:
                    if run.text.strip() and contains_meaningful_content(run.text):
                        extracted_id = run.text.strip() # get id

                        try:           
                            extracted_id = int(extracted_id) # convert
                            run.text = translation_hashmap.get(extracted_id,"")
                        except:
                            run.text = ""


def process_shapes_recursive(shapes):
    """Process shapes to extract text and assign IDs."""
    for shape in shapes:
        if shape.has_text_frame:
            extract_shape_text(shape)
        elif shape.has_table:
            extract_table_text(shape.table)
        elif hasattr(shape, 'shapes'):
            process_shapes_recursive(shape.shapes)

def process_presentation(input_file):
    print(f"Opening {input_file}")
    try:
        input_ppt = Presentation(input_file)
    except Exception as e:
        print(f"Error opening file {input_file}: {e}")
        return []

    for slide in input_ppt.slides:
        process_shapes_recursive(slide.shapes)

    # rename the file to take out the all the weird numbers
    output_file = knio.flow_variables["tempPPTXpath_parent"] + "/" + knio.flow_variables["cleanedFileName"]
    
    # save the updated powerpoint to the temp file with text replaced by their unique id
    input_ppt.save(output_file)
    print("Saving translated powerpoint")

def main():
    process_presentation(input_file)
    knio.output_tables[0] = knio.input_tables[0] # so knime doesnt complain about not outputting a table

main()
