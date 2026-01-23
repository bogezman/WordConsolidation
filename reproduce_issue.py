import zipfile
import io
import re
import unittest

def extract_authors(uploaded_file):
    """
    Copy of the function from app.py to test logic in isolation.
    """
    authors = set()
    
    # Regex to find authors
    pattern_attr_double = re.compile(rb'((?:w|w15):author=")([^"]*)(")')
    pattern_attr_single = re.compile(rb"((?:w|w15):author=')([^']*)(')")
    
    # Element-text based
    pattern_el_creator = re.compile(rb'(<dc:creator>)(.*?)(</dc:creator>)')
    pattern_el_lastmod = re.compile(rb'(<cp:lastModifiedBy>)(.*?)(</cp:lastModifiedBy>)')

    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            for item in zin.infolist():
                if item.filename.endswith('.xml'):
                    content = zin.read(item.filename)
                    
                    # Attribute scan
                    for match in pattern_attr_double.finditer(content):
                        authors.add(match.group(2).decode('utf-8').strip())
                    for match in pattern_attr_single.finditer(content):
                        authors.add(match.group(2).decode('utf-8').strip())
                    
                    # Element scan
                    for match in pattern_el_creator.finditer(content):
                         authors.add(match.group(2).decode('utf-8').strip())
                    for match in pattern_el_lastmod.finditer(content):
                         authors.add(match.group(2).decode('utf-8').strip())
                         
    except Exception:
        pass
        
    return sorted(list(authors))

class TestDuplicateAuthors(unittest.TestCase):
    def test_whitespace_duplication(self):
        # Create a dummy zip/docx in memory
        input_buffer = io.BytesIO()
        with zipfile.ZipFile(input_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            # Simulate one author "Test User" and another "Test User " (with space)
            # Word sometimes does this or user input error.
            
            # document.xml has "Test User"
            xml_content_1 = b'<?xml version="1.0"?><w:document><w:comments><w:comment w:id="0" w:author="Test User" ...></w:comment></w:comments></w:document>'
            z.writestr('word/document.xml', xml_content_1)
            
            # core.xml has "Test User " (with trailing space)
            xml_content_2 = b'<?xml version="1.0"?><cp:coreProperties><dc:creator>Test User </dc:creator></cp:coreProperties>'
            z.writestr('docProps/core.xml', xml_content_2)

        input_buffer.seek(0)
        
        extracted = extract_authors(input_buffer)
        print(f"Extracted authors: {extracted}")
        
        # We expect one entry if whitespace is stripped
        self.assertIn("Test User", extracted)
        self.assertEqual(len(extracted), 1, "Should have 1 unique author if whitespace is stripped")

if __name__ == '__main__':
    unittest.main()
