import streamlit as st
import zipfile
import io
import re
import os

def process_docx(uploaded_file, new_author_name, new_initials):
    """
    Reads a docx file (as a zip), modifies XML content in memory to replace author names and initials,
    and returns a bytes object of the new docx file.
    """
    # Create a buffer for the new docx
    output_buffer = io.BytesIO()
    
    # regex patterns for replacing author and initials
    # We look for w:author="Value" and w:initials="Value"
    # Using byte strings for regex since we read files as bytes
    
    # Build regex patterns handling both single and double quotes
    # w:author
    pattern_author_double = re.compile(rb'(w:author=")([^"]*)(")')
    pattern_author_single = re.compile(rb"(w:author=')([^']*)(')")
    
    # w:initials
    pattern_initials_double = re.compile(rb'(w:initials=")([^"]*)(")')
    pattern_initials_single = re.compile(rb"(w:initials=')([^']*)(')")

    # Convert new values to bytes (utf-8)
    new_author_bytes = new_author_name.encode('utf-8')
    new_initials_bytes = new_initials.encode('utf-8')

    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    
                    # We only want to modify XML files that might contain author info.
                    # Usually these are word/document.xml, word/comments.xml, word/settings.xml, etc.
                    # To be safe and comprehensive, we can check typical xml files or just all .xml files.
                    if item.filename.endswith('.xml'):
                        # Apply replacements
                        # Replace author
                        content = pattern_author_double.sub(rb'\g<1>' + new_author_bytes + rb'\g<3>', content)
                        content = pattern_author_single.sub(rb'\g<1>' + new_author_bytes + rb'\g<3>', content)
                        
                        # Replace initials
                        content = pattern_initials_double.sub(rb'\g<1>' + new_initials_bytes + rb'\g<3>', content)
                        content = pattern_initials_single.sub(rb'\g<1>' + new_initials_bytes + rb'\g<3>', content)
                    
                    # Write content (modified or original) to the new zip
                    zout.writestr(item, content)
                    
        return output_buffer.getvalue()

    except zipfile.BadZipFile:
        st.error("Error: The uploaded file is not a valid docx or zip file.")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        return None

def main():
    st.set_page_config(page_title="WordConsolidation", page_icon="📝")
    
    st.title("WordConsolidation 📝")
    st.markdown("""
    Upload a `.docx` file to automatically replace all revision authors and comment authors with your specified name.
    
    **Privacy Note:** All processing is done in-memory. Your files are not saved to the server.
    """)
    
    # Sidebar for inputs
    with st.sidebar:
        st.header("Configuration")
        new_name = st.text_input("New Author Name", value="Reviewer")
        new_initials = st.text_input("New Initials", value="REV")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose a Word Document", type=["docx"])
    
    if uploaded_file is not None:
        # Show file details
        file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": f"{uploaded_file.size / 1024:.2f} KB"}
        st.write(file_details)
        
        # Process button
        if st.button("Process Document"):
            with st.spinner("Processing document..."):
                processed_data = process_docx(uploaded_file, new_name, new_initials)
                
            if processed_data:
                st.success("Processing complete!")
                
                # Create a download button
                new_filename = f"consolidated_{uploaded_file.name}"
                st.download_button(
                    label="Download Consolidated Document",
                    data=processed_data,
                    file_name=new_filename,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

if __name__ == '__main__':
    main()
