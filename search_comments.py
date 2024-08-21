import zipfile
import lxml.etree as ET

word_docx_path = ""
search_term = ""

def search_comments_in_docx(doc_path, search_term):
    # Open the .docx file as a zip
    with zipfile.ZipFile(doc_path, 'r') as docx_zip:
        # Read the main document XML and the comments XML
        document_xml = docx_zip.read('word/document.xml')
        comments_xml = docx_zip.read('word/comments.xml')
        
        # Parse the XML content
        doc_root = ET.fromstring(document_xml)
        comments_root = ET.fromstring(comments_xml)
        
        # Namespace
        ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        
        # Create a dictionary to store comments by their ID
        comments_dict = {}
        for comment in comments_root.findall('.//w:comment', ns):
            comment_id = comment.get(f'{{{ns["w"]}}}id')
            comment_text = "".join(t.text for t in comment.findall('.//w:t', ns)).strip()
            comments_dict[comment_id] = comment_text
        
        # Find all the commented texts in the document
        for comment_start in doc_root.findall('.//w:commentRangeStart', ns):
            comment_id = comment_start.get(f'{{{ns["w"]}}}id')
            commented_text = ""
            
            # Navigate through the document to find the text associated with this comment
            next_element = comment_start.getnext()
            while next_element is not None and not next_element.tag.endswith('commentRangeEnd'):
                if next_element.tag.endswith('r'):
                    commented_text += "".join(t.text for t in next_element.findall('.//w:t', ns) if t.text)
                next_element = next_element.getnext()
            
            # Check if the comment text contains the search term
            if comment_id in comments_dict and search_term.lower() in comments_dict[comment_id].lower():
                print(f"\"{commented_text}\": \"{comments_dict[comment_id]}\"")


search_comments_in_docx(word_docx_path, search_term)
