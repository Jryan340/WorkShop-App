from ebooklib import epub
import ebooklib  # Ensure ebooklib is imported
import re
import openpyxl

# Function to format the chapter name based on item ID patterns
def format_chapter_name(item_id):
    # Define patterns and associated chapter names
    if re.search(r'\bded\b', item_id, re.IGNORECASE):
        return "Dedication"
    elif re.search(r'\back\b', item_id, re.IGNORECASE):
        return "Acknowledgements"
    elif re.search(r'\bpre\b', item_id, re.IGNORECASE):
        return "Prelude"
    elif re.search(r'\bpro\b', item_id, re.IGNORECASE):
        return "Prologue"
    elif match := re.search(r'\bc(\d+)\b', item_id, re.IGNORECASE):
        return f"Chapter {int(match.group(1))}"  # Convert to integer for clean display
    elif match := re.search(r'\bp(\d+)\b', item_id, re.IGNORECASE):
        return f"Interlude {int(match.group(1))}"  # Interludes based on p and number pattern
    else:
        # Remove '.xhtml' extension if present, and return the raw item_id as a fallback
        return item_id.replace('.xhtml', '')

# Function to extract and format text from the EPUB file in spine order
def extract_text_from_epub(file_path):
    book = epub.read_epub(file_path)
    chapters = []

    # Loop through spine items in order
    for item_id, _ in book.spine:
        item = book.get_item_with_id(item_id)  # Get the actual item using its ID

        # Process only document items (typically chapters)
        if item and item.get_type() == ebooklib.ITEM_DOCUMENT:
            # Extract text, remove HTML tags, and replace '&#13;' with a newline
            text = re.sub('<[^<]+?>', '', item.get_body_content().decode('utf-8'))
            text = text.replace("&#13;", "\n")  # Replace &#13; with newline
            
            # Format chapter name based on item ID
            chapter_name = format_chapter_name(item_id)
            
            # Append the item ID, formatted chapter name, and content to the chapters list
            chapters.append((item_id, chapter_name, text))

    return chapters

# Function to create an Excel file with item ID, chapter names, and content
def create_excel_with_chapters(chapters, output_path="epub_content_v2.xlsx"):
    # Initialize a new Excel workbook and select the active worksheet
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "EPUB Content"

    # Add headers
    sheet["A1"] = "Item ID"
    sheet["B1"] = "Chapter Name"
    sheet["C1"] = "Content"

    # Write each chapter's content to the Excel sheet
    for item_id, chapter_name, content in chapters:
        # Append item ID, chapter name, and content to the Excel sheet
        sheet.append([item_id, chapter_name, content])
        
    # Save the workbook
    workbook.save(output_path)
    print(f"Excel file saved as '{output_path}'")

# Define file paths
epub_file_path = "C:/Users/jryan/Documents/python/The_Way_of_Kings.epub"
output_excel_path = "C:/Users/jryan/Documents/python/The_Way_of_Kings.xlsx"

# Extract text from EPUB and create the Excel file
chapters = extract_text_from_epub(epub_file_path)
create_excel_with_chapters(chapters, output_excel_path)
