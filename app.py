import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io, re, zipfile, os, datetime
from supabase import create_client, Client

# Register font (ensure this font file exists in your working folder)
# NOTE: Using a relative path './Bliss Extra Bold.ttf' requires the font file to be present.
# Ensure font file 'Bliss Extra Bold.ttf' is present!
try:
    pdfmetrics.registerFont(TTFont('BlissExtraBold', './Bliss Extra Bold.ttf'))
except Exception as e:
    st.warning(f"Warning: Could not register custom font. Ensure 'Bliss Extra Bold.ttf' is in the root directory. Using default fonts for now. Error: {e}")

# Supabase setup (use Streamlit secrets for credentials)
SUPABASE_URL = st.secrets["supabase"]["url"]
SUPABASE_KEY = st.secrets["supabase"]["key"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# Database setup
def init_db():
    pass 

init_db()

# Streamlit page setup
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")

# File uploads - Only Excel now
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])

# Fixed certificate template path
template_path = './certificate_template.pdf'
if not os.path.exists(template_path):
    st.error("Certificate template not found. Please ensure 'certificate_template.pdf' is in the working directory.")
    st.stop()

def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple (0-1 range for ReportLab)."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) / 255 for i in (0, 2, 4))

# Preset certificate text settings
student_font_size = 18
student_x = 427
student_y = 200
student_font = "Helvetica-Bold"
student_color = "#000000"

school_font_size = 18
school_x = 306
school_y = 550
school_font = "Helvetica-Bold"
school_color = "#000000"

# User input fields for database storage
st.markdown("---")
st.markdown(" User Details ")
user_name = st.text_input("Name", "")
user_school_name = st.text_input("School Name", "")
user_school_number = st.text_input("School Number", "")
user_contact_number = st.text_input("Contact Number", "")
user_ic_number = st.text_input("IC Number", "") 

# Validate all required fields (FIXED: Added user_ic_number)
inputs_valid = all([
    user_name.strip(), 
    user_school_name.strip(), 
    user_school_number.strip(), 
    user_contact_number.strip(), 
    user_ic_number.strip()
])

st.markdown("---")

# Main logic
if excel_file and inputs_valid:
    
    # Read Excel and ensure headers are correct
    participants = pd.read_excel(excel_file, header=0)
    
    # Handle auto-generated column names if the first cell is empty
    if participants.columns[0] == "" or participants.columns[0] is None:
        participants.columns = ["Student Name"]

    # Auto-detect columns
    student_col = next(
        (c for c in participants.columns if "student" in c.lower() or "name" in c.lower()),
        participants.columns[0]
    )

    school_col = next(
        (c for c in participants.columns if "school" in c.lower() or "institution" in c.lower()),
        None if len(participants.columns) == 1 else participants.columns[1]
    )

    st.markdown(f"**Detected Columns:** Using **`{student_col}`** for Student names.")
    st.markdown(f"Using **`{school_col}`** for School names.")
    st.info(f"Ready to process **{len(participants)}** participants.")

    # Generate button
    if st.button("Generate Certificates and Download Zip"):
        
        # --- START STATUS MESSAGE ---
        st.info("Starting generation process... Please wait.")
        
        # --- END STATUS MESSAGE ---
        
        current_time = datetime.datetime.now(datetime.timezone.utc).isoformat()
        
        # Store user details in Supabase database
        data = {
            "name": user_name.strip(),
            "school_name": user_school_name.strip(),
            "school_number": user_school_number.strip(),
            "contact_number": user_contact_number.strip(),
            "ic_number": user_ic_number.strip(), 
            "timestamp": current_time,
        }
        
        # --- DATABASE INSERTION BLOCK ---
        try:
            response = supabase.table("Generations").insert(data).execute()
            
            if response.data:
                st.success("User details stored successfully in the database.")
            elif response.error:
                st.error(f"Failed to store data in database. Supabase Error: {response.error['message']}")
            else:
                st.error("Failed to store data in database. Response data was empty.")

        except Exception as e:
            st.error(f"Database insertion failed due to an unexpected error: {e}")
            # If database logging is critical, uncomment the line below:
            # st.stop() 
        # --- END DATABASE INSERTION BLOCK ---

        
        zip_buf = io.BytesIO()
        success, fail = 0, 0

        # --- CERTIFICATE GENERATION BLOCK ---
        st.caption("Generating PDFs and creating ZIP file...")
        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            status_placeholder = st.empty() 
            
            for idx, row in participants.iterrows():
                status_placeholder.info(f"Processing certificate {idx + 1} of {len(participants)} for: {row[student_col]}...")
                
                try:
                    # Validate student name
                    raw_student = row[student_col]
                    if pd.isna(raw_student) or str(raw_student).strip() == "":
                        fail += 1
                        continue
                    student = str(raw_student).strip()

                    # Optional school name
                    school = ""
                    if school_col is not None:
                        raw_school = row[school_col]
                        if not pd.isna(raw_school) and str(raw_school).strip() != "":
                            school = str(raw_school).strip()

                    # PDF Logic (Read, Overlay, Merge, Write)
                    base_reader = PdfReader(template_path)
                    base_page = base_reader.pages[0]
                    media_box = base_page.mediabox
                    w = float(media_box.width)
                    h = float(media_box.height)

                    overlay_packet = io.BytesIO()
                    c = canvas.Canvas(overlay_packet, pagesize=(w, h))
                    
                    # Draw Student Name
                    c.setFillColorRGB(*hex_to_rgb(student_color))  
                    c.setFont(student_font, student_font_size)     
                    c.drawCentredString(student_x, student_y, student)

                    # Draw School Name
                    if school:
                        c.setFillColorRGB(*hex_to_rgb(school_color))  
                        c.setFont(school_font, school_font_size)      
                        c.drawCentredString(school_x, school_y, school)

                    c.save()
                    overlay_packet.seek(0)

                    overlay_reader = PdfReader(overlay_packet)
                    merged_page = PageObject.create_blank_page(width=w, height=h)
                    merged_page.merge_page(base_page)
                    merged_page.merge_page(overlay_reader.pages[0])

                    out_buf = io.BytesIO()
                    writer = PdfWriter()
                    writer.add_page(merged_page)
                    writer.write(out_buf)
                    out_buf.seek(0)

                    # Safe filename and add to ZIP
                    safe_name = re.sub(r'[<>:"/\\|?*]', "_", student.replace(" ", "_"))
                    filename = f"{idx+1:03d}_{safe_name}_certificate.pdf"

                    zipf.writestr(filename, out_buf.getvalue())
                    success += 1

                except Exception as e:
                    fail += 1
                    status_placeholder.warning(f"Error creating certificate for {student or 'Unknown'} (Row {idx+1}): {e}")

        # Finalize and download
        status_placeholder.empty() # Clear the status update message
        # --- END CERTIFICATE GENERATION BLOCK ---
        
        zip_buf.seek(0)
        
        # --- FINAL SUCCESS MESSAGE AND DOWNLOAD BUTTON ---
        st.balloons()
        st.success(f"Generation Completed! **{success}** successful, **{fail}** failed.")

        st.download_button(
            "Download All Certificates (.zip)",
            data=zip_buf.getvalue(),
            file_name="certificates.zip",
            mime="application/zip"
        )
        # --- END FINAL ---

elif not inputs_valid:
    st.warning("Please fill in **all** required user details (Name, School Name, School Number, Contact Number, IC Number) to proceed.")
elif excel_file is None:
    st.info("Please upload the Excel file and fill in the required user details to begin.")
