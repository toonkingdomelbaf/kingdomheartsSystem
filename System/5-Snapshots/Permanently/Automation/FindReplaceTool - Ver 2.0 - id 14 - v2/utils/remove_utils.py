from docx import Document

def remove_blank_line_at_matches(doc_or_text, matches_list, position="after"):
    """
    Remove **one blank paragraph** per match in a DOCX file either after or before the matched paragraph.
    
    Parameters:
        doc_or_text (str): Path to the DOCX file.
        matches_list (list): List of matches in the format [(index, matched_text, positions), ...].
        position (str): "after" or "before" the matched paragraph.
    
    Returns:
        int: Total blank paragraphs removed.
    """
    removed_count = 0

    try:
        doc = Document(doc_or_text)
        paragraphs = doc.paragraphs

        for match_index, matched_text, _ in matches_list:
            # Find the paragraph that matches the text
            for idx, para in enumerate(paragraphs):
                if para.text.strip() == matched_text.strip():
                    # Determine sibling index
                    sibling_idx = idx + 1 if position == "after" else idx - 1

                    if 0 <= sibling_idx < len(paragraphs):
                        sibling_para = paragraphs[sibling_idx]
                        print(f"DEBUG MATCHED PARAGRAPH: '{para.text}'")
                        print(f"DEBUG POSITION: {position}")
                        print(f"DEBUG CHECKING SIBLING INDEX: {sibling_idx} TAG: {sibling_para._element.tag}")
                        print(f"DEBUG SIBLING TEXT: '{sibling_para.text}'")

                        if sibling_para.text.strip() == "":
                            sibling_para._element.getparent().remove(sibling_para._element)
                            removed_count += 1
                            print("DEBUG REMOVE BLANK PARAGRAPH")
                        else:
                            print("DEBUG STOP: NOT BLANK")
                    break  # Only remove one blank per match

        doc.save(doc_or_text)
        print(f"DEBUG TOTAL REMOVED: {removed_count}")
        return removed_count

    except Exception as e:
        print(f"Error processing file {doc_or_text}: {e}")
        return 0
