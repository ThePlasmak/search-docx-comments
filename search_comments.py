from docx import Document
import lxml.etree as ET

word_docx_path = ""
search_term = ""

def search_comments_in_docx(doc_path, search_term):
    # Load the Word document
    doc = Document(doc_path)

    # Find the comments part by searching for the appropriate relationship type
    comments_part = None
    for rel in doc.part.rels.values():
        if "comments" in rel.target_ref:
            comments_part = rel.target_part
            break
    
    if not comments_part:
        print("No comments found in the document.")
        return
    
    # Parse the comments XML
    comments_xml = comments_part.blob
    root = ET.fromstring(comments_xml)
    
    # Initialize a list to store comments containing the search term
    matching_comments = []
    
    # Define the namespace (handled differently without using xpath)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Find all comments
    for comment in root.findall('.//w:comment', ns):
        # Extract the comment text
        comment_text = "".join(t.text for t in comment.findall('.//w:t', ns)).strip()
        
        # Extract the text being commented on using the comment id
        comment_id = comment.get(f'{{{ns["w"]}}}id')
        comment_ref = doc.element.findall(f'.//w:commentRangeStart[@w:id="{comment_id}"]', namespaces=ns)
        
        if comment_ref:
            commented_text = ""
            next_element = comment_ref[0].getnext()
            while next_element is not None:
                if next_element.tag.endswith('commentRangeEnd'):
                    break
                if next_element.tag.endswith('r'):
                    commented_text += "".join(t.text for t in next_element.findall('.//w:t', ns) if t.text)
                next_element = next_element.getnext()
        
        if search_term.lower() in comment_text.lower():
            matching_comments.append((commented_text, comment_text))
    
    # Print or return the list of matching comments
    if matching_comments:
        print(f"Comments containing \"{search_term}\":\n")
        for commented_text, comment_text in matching_comments:
            print(f"\"{commented_text}\": \"{comment_text}\"")
    else:
        print(f"No comments found containing the term '{search_term}'.")
    
    # Keep the terminal open
    input("\nPress Enter to close...")

search_comments_in_docx(word_docx_path, search_term)
