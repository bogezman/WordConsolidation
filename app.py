import streamlit as st
import zipfile
import io
import re
import os

def extract_authors(uploaded_file):
    """
    Extracts a set of unique authors from the docx file.
    """
    authors = set()
    
    # Regex to find authors
    # 1. Attribute-based: w:author="Value" or w:author='Value' (and w15:author)
    pattern_attr_double = re.compile(rb'((?:w|w15):author=")([^"]*)(")')
    pattern_attr_single = re.compile(rb"((?:w|w15):author=')([^']*)(')")
    
    # 2. Element-text based: <dc:creator>Value</dc:creator> or <cp:lastModifiedBy>Value</cp:lastModifiedBy>
    # Note: These are in docProps/core.xml usually.
    pattern_el_creator = re.compile(rb'(<dc:creator>)(.*?)(</dc:creator>)')
    pattern_el_lastmod = re.compile(rb'(<cp:lastModifiedBy>)(.*?)(</cp:lastModifiedBy>)')

    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            for item in zin.infolist():
                if item.filename.endswith('.xml'):
                    content = zin.read(item.filename)
                    
                    # Attribute scan
                    for match in pattern_attr_double.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
                    for match in pattern_attr_single.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
                    
                    # Element scan
                    for match in pattern_el_creator.finditer(content):
                         authors.add(match.group(2).decode('utf-8'))
                    for match in pattern_el_lastmod.finditer(content):
                         authors.add(match.group(2).decode('utf-8'))
                         
    except Exception:
        pass # Ignore errors during extraction to avoid breaking the flow if zip is bad (will be caught later)
        
    return sorted(list(authors))

def process_docx(uploaded_file, target_authors, new_author_name, new_initials, remove_highlights=False):
    """
    Reads a docx file (as a zip), modifies XML content in memory to replace author names and initials,
    and returns a bytes object of the new docx file.
    """
    # Create a buffer for the new docx
    output_buffer = io.BytesIO()
    
    # regex patterns for replacing author and initials
    # We look for w:author="Value" and w:initials="Value"
    # Using byte strings for regex since we read files as bytes
    
    pattern_author_double = re.compile(rb'((?:w|w15):author=")([^"]*)(")')
    pattern_author_single = re.compile(rb"((?:w|w15):author=')([^']*)(')")
    
    pattern_initials_double = re.compile(rb'(w:initials=")([^"]*)(")')
    pattern_initials_single = re.compile(rb"(w:initials=')([^']*)(')")

    # Metadata element patterns
    pattern_el_creator = re.compile(rb'(<dc:creator>)(.*?)(</dc:creator>)')
    pattern_el_lastmod = re.compile(rb'(<cp:lastModifiedBy>)(.*?)(</cp:lastModifiedBy>)')

    # Highlight patterns
    # Matches <w:highlight ... /> or <w15:highlight ... />
    # We want to remove the entire tag.
    pattern_highlight = re.compile(rb'(<w(?:15)?:highlight[^>]*/>)')

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
                        
                        if remove_highlights:
                            # Remove highlight tags
                            content = pattern_highlight.sub(rb'', content)
                        
                        # Only proceed with author replacement if target_authors is provided
                        if target_authors:
                            # Apply global string replacement for each selected author
                            # This covers: attributes, metadata elements, AND body text/field results.
                            for author in target_authors:
                                author_bytes = author.encode('utf-8')
                                # Simple replace
                                # Note: This replaces ALL occurrences of the author name.
                                if author_bytes in content:
                                    content = content.replace(author_bytes, new_author_bytes)
                            
                            # We still run the regex for INITIALS separately because Initials are NOT in the target_authors list
                            # (target_authors are full names). 
                            # The user wants to replace "Zhang, Lin" (Author) -> "NewName".
                            # Implicitly, they might want to replace Initials too.
                            # Our previous logic replaced ALL initials blindly. We will keep that for safety/anonymization.
                            
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
    
    tab_tool, tab_about = st.tabs(["Tool", "About"])
    
    with tab_tool:
        st.markdown("""
        Upload a `.docx` file to automatically replace all revision authors and comment authors with your specified name.
        
        **Privacy Note:** All processing is done in-memory. Your files are not saved to the server.
        """)
        
        # Sidebar for inputs - NOTE: Sidebars are global, but we can keep the code here or move it out if we want it to persist across tabs.
        # Usually sidebar config is fine to stay global or be defined here.
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
            
            # Extract authors
            all_authors = extract_authors(uploaded_file)
            
            # Author selection
            st.subheader("Select Authors to Modify")
            if not all_authors:
                st.warning("No authors found in the document.")
                target_authors = []
            else:
                target_authors = st.multiselect(
                    "Authors",
                    options=all_authors,
                    default=all_authors,
                    help="Deselect authors you want to keep unchanged."
                )
            
            # Highlight removal option
            remove_highlights = st.checkbox("Clear all highlights", value=False)
            
            # Process button
            if st.button("Process Document"):
                with st.spinner("Processing document..."):
                    processed_data = process_docx(uploaded_file, target_authors, new_name, new_initials, remove_highlights=remove_highlights)
                    
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

    with tab_about:
        st.header("About WordConsolidation")
        st.markdown("""
        **WordConsolidation** is a privacy-focused tool designed to help you sanitize and anonymize Microsoft Word documents before sharing them. 
        
        ### Key Features
        
        #### 1. Author Anonymization 👤
        Automatically replace author names and initials in:
        - Tracked changes (Revisions)
        - Comments
        - Document Metadata (Creator, Last Modified By)
        
        #### 2. Selective Processing 🎯
        You are in control.
        - Scan the document to list all unique authors.
        - **Select exactly which authors** you want to anonymize.
        - Keep specific team members' names if needed, while anonymizing others.
        
        #### 3. Highlight Removal 🖌️
        Clean up your document with a single click.
        - Remove all highlighter formatting from the text.
        - Useful for finalizing documents after review sessions.
        
        #### 4. Privacy First 🔒
        - **No Data Retention:** Files are processed entirely in-memory.
        - **Secure:** Your original files are never saved to our servers. once you close the tab, the data is gone.
        
        ---
        
        ### How to Use
        1. **Upload** your `.docx` file in the **Tool** tab.
        2. **Configure** the new name and initials in the sidebar (default: "Reviewer").
        3. **Select** the authors you wish to replace from the list.
        4. (Optional) Check **Clear all highlights** to remove highlighting.
        5. Click **Process Document** and download your sanitized file.
        
        ---
        *Version 1.1*
        """)
        
if __name__ == '__main__':
    main()
