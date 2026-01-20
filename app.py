import streamlit as st
import zipfile
import io
import re
import os

# Word standard highlight colors with their hex values for UI preview
HIGHLIGHT_COLORS = {
    "yellow": "#FFFF00",
    "green": "#00FF00",
    "cyan": "#00FFFF",
    "magenta": "#FF00FF",
    "blue": "#0000FF",
    "red": "#FF0000",
    "darkBlue": "#000080",
    "darkCyan": "#008080",
    "darkGreen": "#008000",
    "darkMagenta": "#800080",
    "darkRed": "#800000",
    "darkYellow": "#808000",
    "lightGray": "#C0C0C0",
    "darkGray": "#808080",
}

def extract_revision_authors(uploaded_file):
    """
    Extracts unique authors from tracked changes (w:ins and w:del elements).
    These are insertions and deletions, not comments or metadata.
    """
    authors = set()
    
    # Pattern to find w:ins or w:del elements with w:author attribute
    # Matches: <w:ins ... w:author="Name" ...> or <w:del ... w:author="Name" ...>
    pattern_ins_del = re.compile(rb'<w:(ins|del)[^>]*w:author="([^"]*)"[^>]*>')
    
    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            for item in zin.infolist():
                if item.filename.endswith('.xml'):
                    content = zin.read(item.filename)
                    for match in pattern_ins_del.finditer(content):
                        authors.add(match.group(2).decode('utf-8'))
    except Exception:
        pass
        
    return sorted(list(authors))

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

def apply_author_highlights(uploaded_file, author_colors):
    """
    Applies highlight colors to tracked changes (insertions/deletions) by specific authors.
    
    Args:
        uploaded_file: The docx file as a file-like object
        author_colors: Dict mapping author name to highlight color name (e.g., {'John': 'yellow'})
    
    Returns:
        Bytes of the modified docx file, or None on error
    """
    output_buffer = io.BytesIO()
    
    try:
        with zipfile.ZipFile(uploaded_file, 'r') as zin:
            with zipfile.ZipFile(output_buffer, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    content = zin.read(item.filename)
                    
                    if item.filename.endswith('.xml'):
                        # Process each author's revisions
                        for author, color in author_colors.items():
                            author_bytes = author.encode('utf-8')
                            highlight_tag = f'<w:highlight w:val="{color}"/>'.encode('utf-8')
                            
                            # Pattern to find w:ins or w:del elements for this author
                            # We need to add highlight to runs inside these elements
                            # Strategy: Find <w:ins...author="Name"...>...</w:ins> blocks
                            # and inject highlight into their <w:rPr> elements
                            
                            # Match w:ins or w:del opening tag with this author
                            pattern = rb'(<w:(ins|del)[^>]*w:author="' + re.escape(author_bytes) + rb'"[^>]*>)'
                            
                            def inject_highlight(match_obj):
                                """Process the content after matching w:ins/w:del tag"""
                                return match_obj.group(0)
                            
                            # More robust approach: find all runs within ins/del by author
                            # and add highlight to their run properties
                            
                            # Pattern for ins/del block with author
                            ins_pattern = re.compile(
                                rb'(<w:ins[^>]*w:author="' + re.escape(author_bytes) + rb'"[^>]*>)(.*?)(</w:ins>)',
                                re.DOTALL
                            )
                            del_pattern = re.compile(
                                rb'(<w:del[^>]*w:author="' + re.escape(author_bytes) + rb'"[^>]*>)(.*?)(</w:del>)',
                                re.DOTALL
                            )
                            
                            def add_highlight_to_runs(m):
                                """Add highlight to all w:r elements within the matched block"""
                                opening = m.group(1)
                                inner = m.group(2)
                                closing = m.group(3)
                                
                                # Find w:r elements and add highlight to w:rPr
                                # Pattern: <w:r>...<w:rPr>...</w:rPr>...<w:t>...</w:t>...</w:r>
                                # or <w:r>...<w:t>...</w:t>...</w:r> (no rPr)
                                
                                def process_run(run_match):
                                    run_content = run_match.group(0)
                                    
                                    # Remove any existing highlight tag before adding new one
                                    run_content = re.sub(rb'<w:highlight[^>]*/>', b'', run_content)
                                    
                                    # Check if has w:rPr
                                    if b'<w:rPr>' in run_content:
                                        # Insert highlight after <w:rPr>
                                        run_content = run_content.replace(b'<w:rPr>', b'<w:rPr>' + highlight_tag, 1)
                                    elif b'<w:rPr ' in run_content:
                                        # Handle <w:rPr ...> with attributes
                                        run_content = re.sub(
                                            rb'(<w:rPr[^>]*>)',
                                            rb'\1' + highlight_tag,
                                            run_content,
                                            count=1
                                        )
                                    else:
                                        # No w:rPr, need to add one after <w:r> or <w:r ...>
                                        run_content = re.sub(
                                            rb'(<w:r(?:\s[^>]*)?>)',
                                            rb'\1<w:rPr>' + highlight_tag + rb'</w:rPr>',
                                            run_content,
                                            count=1
                                        )
                                    return run_content
                                
                                # Match w:r elements (including w:r with attributes)
                                inner = re.sub(rb'<w:r(?:\s[^>]*)?>.*?</w:r>', process_run, inner, flags=re.DOTALL)
                                
                                return opening + inner + closing
                            
                            content = ins_pattern.sub(add_highlight_to_runs, content)
                            content = del_pattern.sub(add_highlight_to_runs, content)
                    
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
    
    tab_sanitize, tab_highlight_rev, tab_about = st.tabs(["Sanitize", "Highlight Revisions", "About"])
    
    with tab_sanitize:
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

    with tab_highlight_rev:
        st.header("Highlight Author Revisions 🖍️")
        st.info("🧪 **Experimental Feature** - This feature is new and may have limitations with complex documents.")
        st.markdown("""
        Upload a `.docx` file to highlight tracked changes (insertions and deletions) by specific authors.
        
        **How it works:** Select authors and assign them Word-standard highlight colors. 
        The tool will apply highlighting to all their additions and deletions.
        """)
        
        # File uploader for highlight tab
        highlight_file = st.file_uploader("Choose a Word Document", type=["docx"], key="highlight_uploader")
        
        if highlight_file is not None:
            # Show file details
            file_details = {"FileName": highlight_file.name, "FileType": highlight_file.type, "FileSize": f"{highlight_file.size / 1024:.2f} KB"}
            st.write(file_details)
            
            # Extract revision authors
            revision_authors = extract_revision_authors(highlight_file)
            
            if not revision_authors:
                st.warning("No tracked changes found in this document. Make sure the document has revisions (insertions/deletions).")
            else:
                st.subheader("Select Authors & Assign Colors")
                
                # Create color options with preview
                color_options = list(HIGHLIGHT_COLORS.keys())
                
                # Store author-color mappings
                author_color_selections = {}
                
                # Create a grid of author selections with color pickers
                for idx, author in enumerate(revision_authors):
                    col1, col2, col3, col4 = st.columns([0.5, 2, 2.5, 0.5])
                    
                    with col1:
                        # Checkbox to include this author
                        include = st.checkbox("", value=True, key=f"include_{idx}")
                    
                    with col2:
                        st.write(f"**{author}**")
                    
                    with col3:
                        if include:
                            # Color selector with preview
                            color_name = st.selectbox(
                                "Color",
                                options=color_options,
                                index=idx % len(color_options),  # Rotate default colors
                                key=f"color_{idx}",
                                label_visibility="collapsed"
                            )
                            author_color_selections[author] = color_name
                    
                    with col4:
                        if include:
                            # Show color preview inline
                            hex_color = HIGHLIGHT_COLORS[color_name]
                            st.markdown(
                                f'<div style="width:100%;height:38px;background-color:{hex_color};border-radius:4px;border:1px solid #ccc;margin-top:0px;"></div>',
                                unsafe_allow_html=True
                            )
                
                # Color legend
                with st.expander("View All Available Colors"):
                    legend_cols = st.columns(4)
                    for i, (name, hex_val) in enumerate(HIGHLIGHT_COLORS.items()):
                        with legend_cols[i % 4]:
                            st.markdown(
                                f'<div style="display:flex;align-items:center;margin:4px 0;">'
                                f'<div style="width:20px;height:20px;background-color:{hex_val};border-radius:3px;border:1px solid #ccc;margin-right:8px;"></div>'
                                f'<span style="font-size:12px;">{name}</span></div>',
                                unsafe_allow_html=True
                            )
                
                # Process button
                if st.button("Apply Highlights", key="apply_highlights_btn"):
                    if not author_color_selections:
                        st.warning("Please select at least one author to highlight.")
                    else:
                        with st.spinner("Applying highlights..."):
                            # Reset file position
                            highlight_file.seek(0)
                            processed_data = apply_author_highlights(highlight_file, author_color_selections)
                        
                        if processed_data:
                            st.success("Highlights applied successfully!")
                            
                            # Preview what was done
                            st.write("**Applied highlights:**")
                            for author, color in author_color_selections.items():
                                hex_val = HIGHLIGHT_COLORS[color]
                                st.markdown(
                                    f'<div style="display:flex;align-items:center;margin:4px 0;">'
                                    f'<div style="width:16px;height:16px;background-color:{hex_val};border-radius:2px;border:1px solid #ccc;margin-right:8px;"></div>'
                                    f'<span>{author}</span></div>',
                                    unsafe_allow_html=True
                                )
                            
                            # Download button
                            new_filename = f"highlighted_{highlight_file.name}"
                            st.download_button(
                                label="Download Highlighted Document",
                                data=processed_data,
                                file_name=new_filename,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                key="download_highlighted"
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
        
        #### 4. Revision Highlighting 🖍️ *(NEW)*
        Visually identify changes by specific authors.
        - Upload a document with tracked changes.
        - Select authors and assign Word-standard highlight colors.
        - Additions and deletions by those authors get highlighted.
        - Color preview swatches help you pick the right color.
        
        #### 5. Privacy First 🔒
        - **No Data Retention:** Files are processed entirely in-memory.
        - **Secure:** Your original files are never saved to our servers. once you close the tab, the data is gone.
        
        ---
        
        ### How to Use
        
        **Sanitize Tab** (Anonymization & Cleanup):
        1. **Upload** your `.docx` file in the **Sanitize** tab.
        2. **Configure** the new name and initials in the sidebar (default: "Reviewer").
        3. **Select** the authors you wish to replace from the list.
        4. (Optional) Check **Clear all highlights** to remove highlighting.
        5. Click **Process Document** and download your sanitized file.
        
        **Highlight Revisions Tab** (Revision Highlighting):
        1. **Upload** your `.docx` file in the **Highlight Revisions** tab.
        2. **Select authors** and assign colors using the dropdowns.
        3. Click **Apply Highlights** and download the highlighted file.
        
        ---
        *Version 1.2*
        """)
        
if __name__ == '__main__':
    main()
