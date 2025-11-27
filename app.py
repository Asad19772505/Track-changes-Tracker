import streamlit as st
from docx import Document
from lxml import etree
from io import BytesIO
import zipfile
import base64

# ------------------------------
# Helper: Extract document.xml from docx
# ------------------------------
def extract_document_xml(uploaded_file):
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        xml_content = zip_ref.read("word/document.xml")
    return xml_content

# ------------------------------
# Helper: Extract comments.xml (if exists)
# ------------------------------
def extract_comments_xml(uploaded_file):
    with zipfile.ZipFile(uploaded_file, "r") as zip_ref:
        try:
            xml_content = zip_ref.read("word/comments.xml")
            return xml_content
        except KeyError:
            return None  # No comments found

# ------------------------------
# Helper: Show tracked changes
# ------------------------------
def parse_tracked_changes(xml_content):
    root = etree.fromstring(xml_content)

    ins = root.findall(".//w:ins", namespaces=root.nsmap)
    deletes = root.findall(".//w:del", namespaces=root.nsmap)

    insertions = []
    deletions = []

    for i in ins:
        txt = "".join(i.itertext())
        insertions.append(txt)

    for d in deletes:
        txt = "".join(d.itertext())
        deletions.append(txt)

    return insertions, deletions

# ------------------------------
# Helper: Accept or reject all changes
# ------------------------------
def modify_document(uploaded_file, action="accept"):
    memfile = BytesIO()
    
    # Load entire .docx archive
    with zipfile.ZipFile(uploaded_file, "r") as zin:
        with zipfile.ZipFile(memfile, "w") as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                # Modify document.xml only
                if item.filename == "word/document.xml":
                    xml_root = etree.fromstring(data)

                    if action == "accept":
                        # Accept: keep insertions, remove deletions
                        for d in xml_root.findall(".//w:del", xml_root.nsmap):
                            d.getparent().remove(d)
                        for ins in xml_root.findall(".//w:ins", xml_root.nsmap):
                            ins.tag = "{%s}r" % xml_root.nsmap["w"]  # Convert to normal run

                    elif action == "reject":
                        # Reject: remove insertions, keep original text before changes
                        for ins in xml_root.findall(".//w:ins", xml_root.nsmap):
                            ins.getparent().remove(ins)
                        for d in xml_root.findall(".//w:del", xml_root.nsmap):
                            d.tag = "{%s}r" % xml_root.nsmap["w"]

                    data = etree.tostring(xml_root)
                
                zout.writestr(item, data)

    memfile.seek(0)
    return memfile


# ------------------------------
# Streamlit UI
# ------------------------------
st.title("📄 Tracked Changes Analyzer for Word Documents")

uploaded_file = st.file_uploader("Upload a .docx file", type=["docx"])

if uploaded_file:
    st.subheader("🔍 Extracting Tracked Changes...")

    xml_content = extract_document_xml(uploaded_file)
    comments_xml = extract_comments_xml(uploaded_file)

    insertions, deletions = parse_tracked_changes(xml_content)

    st.write("### ✏️ Insertions Found")
    if insertions:
        for ins in insertions:
            st.write(f"- {ins}")
    else:
        st.info("No insertions found.")

    st.write("### ❌ Deletions Found")
    if deletions:
        for d in deletions:
            st.write(f"- {d}")
    else:
        st.info("No deletions found.")

    st.write("### 💬 Comments")
    if comments_xml:
        root = etree.fromstring(comments_xml)
        comments = root.findall(".//w:comment", root.nsmap)
        for c in comments:
            st.write(f"- {''.join(c.itertext())}")
    else:
        st.info("No comments found.")

    st.divider()

    # ----------------------
    # Accept / Reject Buttons
    # ----------------------
    st.subheader("⚙ Process Document")

    col1, col2 = st.columns(2)

    with col1:
        if st.button("Accept All Changes"):
            corrected_doc = modify_document(uploaded_file, action="accept")
            st.success("All changes accepted!")
            st.download_button(
                "Download Clean Document",
                data=corrected_doc,
                file_name="accepted_changes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

    with col2:
        if st.button("Reject All Changes"):
            corrected_doc = modify_document(uploaded_file, action="reject")
            st.error("All insertions removed. Deletions restored!")
            st.download_button(
                "Download Rejected Version",
                data=corrected_doc,
                file_name="rejected_changes.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
