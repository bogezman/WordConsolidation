import zipfile
import io
import re
import unittest

# Copied logic for testing purposes to avoid importing app (and triggering streamlit import)

def extract_authors_logic(uploaded_file):
    authors = set()
    pattern_attr_double = re.compile(rb'((?:w|w15):author=")([^"]*)(")')
    pattern_attr_single = re.compile(rb"((?:w|w15):author=')([^']*)(')")
    pattern_el_creator = re.compile(rb'(<dc:creator>)(.*?)(</dc:creator>)')
    pattern_el_lastmod = re.compile(rb'(<cp:lastModifiedBy>)(.*?)(</cp:lastModifiedBy>)')

    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            for item in zin.infolist():
                if item.filename.endswith('.xml'):
                    content = zin.read(item.filename)
                    for match in pattern_attr_double.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
                    for match in pattern_attr_single.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
                    for match in pattern_el_creator.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
                    for match in pattern_el_lastmod.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
    except Exception:
        pass
    return sorted(list(authors))

def process_docx_logic(uploaded_file, target_authors, new_author_name, new_initials, remove_highlights=False):
    output_buffer = io.BytesIO()
    
    pattern_author_double = re.compile(rb'((?:w|w15):author=")([^"]*)(")')
    pattern_author_single = re.compile(rb"((?:w|w15):author=')([^']*)(')")
    pattern_initials_double = re.compile(rb'(w:initials=")([^"]*)(")')
    pattern_initials_single = re.compile(rb"(w:initials=')([^']*)(')")
    pattern_el_creator = re.compile(rb'(<dc:creator>)(.*?)(</dc:creator>)')
    pattern_el_lastmod = re.compile(rb'(<cp:lastModifiedBy>)(.*?)(</cp:lastModifiedBy>)')
    pattern_highlight = re.compile(rb'(<w(?:15)?:highlight[^>]*/>)')

    new_author_bytes = new_author_name.encode('utf-8')
    new_initials_bytes = new_initials.encode('utf-8')

    with zipfile.ZipFile(uploaded_file, 'r') as zin:
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                if item.filename.endswith('.xml'):
                    if remove_highlights:
                        content = pattern_highlight.sub(rb'', content)

                    # Global string replace for selected authors
                    for author in target_authors:
                        author_bytes = author.encode('utf-8')
                        if author_bytes in content:
                            content = content.replace(author_bytes, new_author_bytes)

                    content = pattern_initials_double.sub(rb'\g<1>' + new_initials_bytes + rb'\g<3>', content)
                    content = pattern_initials_single.sub(rb'\g<1>' + new_initials_bytes + rb'\g<3>', content)
                
                zout.writestr(item, content)
                
    return output_buffer.getvalue()

class TestDocxProcessing(unittest.TestCase):
    def test_replacement(self):
        # Create a dummy zip/docx in memory
        input_buffer = io.BytesIO()
        with zipfile.ZipFile(input_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            # Add a dummy xml file
            xml_content = b'<?xml version="1.0"?><w:document><w:body><w:p w:rsidR="00000000" w:rsidRDefault="00000000"><w:commentRangeStart w:id="0"/><w:r><w:t>Test</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:comments><w:comment w:id="0" w:author="Old Author" w:initials="OA" w:date="2021-01-01T00:00:00Z"><w:p><w:r><w:t>Comment</w:t></w:r></w:p></w:comment></w:comments></w:document>'
            z.writestr('word/document.xml', xml_content)
            # Add a non-xml file
            z.writestr('word/media/image.png', b'fakeimagecontent')

        input_buffer.seek(0)
        
        input_buffer.seek(0)
        
        # Test Extraction
        extracted_authors = extract_authors_logic(input_buffer)
        self.assertIn("Old Author", extracted_authors)
        
        # Reset buffer for reading
        input_buffer.seek(0)

        # Process with partial selection
        # We need another author to test partial selection
        # Re-create zip with two authors
        input_buffer = io.BytesIO()
        with zipfile.ZipFile(input_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            xml_content = b'<?xml version="1.0"?><w:document><w:comments><w:comment w:id="0" w:author="Old Author" w:initials="OA" ...><w:p>...</w:p></w:comment><w:comment w:id="1" w:author="Keep Me" w:initials="KM" ...><w:p>...</w:p></w:comment></w:comments></w:document>'
            z.writestr('word/document.xml', xml_content)
            
            # Add metadata file
            meta_content = b'<?xml version="1.0"?><cp:coreProperties><dc:creator>Creator Name</dc:creator><cp:lastModifiedBy>Modifier Name</cp:lastModifiedBy></cp:coreProperties>'
            z.writestr('docProps/core.xml', meta_content)
            
            # Add footer file with plain text field result
            footer_content = b'<?xml version="1.0"?><w:ftr><w:p><w:r><w:t>Old Author</w:t></w:r></w:p></w:ftr>'
            z.writestr('word/footer1.xml', footer_content)
            
            # Add a non-xml file to verify preservation
            # Add a non-xml file to verify preservation
            z.writestr('word/media/image.png', b'fakeimagecontent')
        
        input_buffer.seek(0)
        
        # Process ONLY "Old Author", "Creator Name", "Modifier Name"
        target_authors = ["Old Author", "Creator Name", "Modifier Name"]
        output_bytes = process_docx_logic(input_buffer, target_authors, "NewName", "NN")
        
        # Verify
        output_buffer = io.BytesIO(output_bytes)
        with zipfile.ZipFile(output_buffer, 'r') as z:
            # Check XML content
            new_xml = z.read('word/document.xml')
            
            # "Old Author" should be replaced
            self.assertIn(b'w:author="NewName"', new_xml)
            self.assertNotIn(b'Old Author', new_xml)
            
            # "Keep Me" should be PRESERVED
            self.assertIn(b'w:author="Keep Me"', new_xml) 

            # Check Metadata
            new_meta = z.read('docProps/core.xml')
            self.assertIn(b'<dc:creator>NewName</dc:creator>', new_meta)
            self.assertIn(b'<cp:lastModifiedBy>NewName</cp:lastModifiedBy>', new_meta)
            self.assertNotIn(b'Creator Name', new_meta)
            self.assertNotIn(b'Modifier Name', new_meta)
            
            # Check Footer (Field Result)
            new_footer = z.read('word/footer1.xml')
            self.assertIn(b'<w:t>NewName</w:t>', new_footer)
            self.assertNotIn(b'Old Author', new_footer)
            
            # Initials are still global replace in current logic
            self.assertIn(b'w:initials="NN"', new_xml)
            # self.assertNotIn(b'OA', new_xml) # Logic replaces all initials
            
            
            # Check non-XML content preserved
            image = z.read('word/media/image.png')
            self.assertEqual(image, b'fakeimagecontent')

        print("Test passed: Author and initials replaced successfully.")

    def test_remove_highlights(self):
        input_buffer = io.BytesIO()
        with zipfile.ZipFile(input_buffer, 'w', zipfile.ZIP_DEFLATED) as z:
            # Add xml content with highlights
            xml_content = b'<?xml version="1.0"?><w:document><w:p><w:r><w:highlight w:val="yellow"/><w:t>Highlighted Text</w:t></w:r></w:p></w:document>'
            z.writestr('word/document.xml', xml_content)
        
        input_buffer.seek(0)
        
        # Process with remove_highlights=True
        output_bytes = process_docx_logic(input_buffer, [], "", "", remove_highlights=True)
        
        output_buffer = io.BytesIO(output_bytes)
        with zipfile.ZipFile(output_buffer, 'r') as z:
            new_xml = z.read('word/document.xml')
            # Ensure highlight tag is gone
            self.assertNotIn(b'w:highlight', new_xml)
            # Ensure text remains
            self.assertIn(b'Highlighted Text', new_xml)

        print("Test passed: Highlights removed successfully.")

if __name__ == '__main__':
    unittest.main()
