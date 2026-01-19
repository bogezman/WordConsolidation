import zipfile
import io
import re
import unittest

# Copied logic for testing purposes to avoid importing app (and triggering streamlit import)
def process_docx_logic(uploaded_file, new_author_name, new_initials):
    output_buffer = io.BytesIO()
    
    pattern_author_double = re.compile(rb'(w:author=")([^"]*)(")')
    pattern_author_single = re.compile(rb"(w:author=')([^']*)(')")
    pattern_initials_double = re.compile(rb'(w:initials=")([^"]*)(")')
    pattern_initials_single = re.compile(rb"(w:initials=')([^']*)(')")

    new_author_bytes = new_author_name.encode('utf-8')
    new_initials_bytes = new_initials.encode('utf-8')

    with zipfile.ZipFile(uploaded_file, 'r') as zin:
        with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                content = zin.read(item.filename)
                
                if item.filename.endswith('.xml'):
                    content = pattern_author_double.sub(rb'\g<1>' + new_author_bytes + rb'\g<3>', content)
                    content = pattern_author_single.sub(rb'\g<1>' + new_author_bytes + rb'\g<3>', content)
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
        
        # Process
        output_bytes = process_docx_logic(input_buffer, "NewName", "NN")
        
        # Verify
        output_buffer = io.BytesIO(output_bytes)
        with zipfile.ZipFile(output_buffer, 'r') as z:
            # Check XML content
            new_xml = z.read('word/document.xml')
            self.assertIn(b'w:author="NewName"', new_xml)
            self.assertIn(b'w:initials="NN"', new_xml)
            self.assertNotIn(b'Old Author', new_xml)
            self.assertNotIn(b'OA', new_xml)
            
            # Check non-XML content preserved
            image = z.read('word/media/image.png')
            self.assertEqual(image, b'fakeimagecontent')

        print("Test passed: Author and initials replaced successfully.")

if __name__ == '__main__':
    unittest.main()
