import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io, re, zipfile, os
from supabase import create_client, Client
import datetime

# Register font (ensure this font file exists in your working folder)

# Supabase setup (use Streamlit secrets for credentials)
SUPABASE_URL = st.secrets["supabase"]["url"]
SUPABASE_KEY = st.secrets["supabase"]["key"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# Database setup (create table if not exists - note: in Supabase, create via dashboard or SQL)
def init_db():
    # Assuming the table 'Generations' is created in Supabase with columns: 
    # id (auto-increment), name, school_name, school_number, contact_number, timestamp, ic_number (MUST BE PRESENT)
    pass 

init_db()

# Streamlit page setup
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")

# File uploads - Only Excel now
excel_file = st.file_uploader("Upload Participant List (Excel)", type=["xlsx"])

# Fixed certificate template path (ensure this file exists in your working folder)
template_path = './certificate_template.pdf'
if not os.path.exists(template_path):
    st.error("Certificate template not found. Please ensure 'certificate_template.pdf' is in the working directory.")
    st.stop()

def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple (0-1 range for ReportLab)."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) / 255 for i in (0, 2, 4))

# Preset certificate text settings (hardcoded, not shown to user)
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
st.markdown("User Details (Required for Database)")
user_name = st.text_input("Name", "")
user_school_name = st.text_input("School Name", "")
user_school_number = st.text_input("School Number", "")
user_contact_number = st.text_input("Contact Number", "")
# New field - must exist in Supabase table
user_ic_number = st.text_input("IC Number", "") 

# Validate required fields
inputs_valid = all([user_name.strip(), user_school_name.strip(), user_school_number.strip(), user_contact_number.strip(), user_ic_number.strip()])

# Main logic
if excel_file and inputs_valid:
    # Read Excel and ensure headers are correct
    participants = pd.read_excel(excel_file, header=0)
    
    # Simple check for empty first column name and rename
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

    st.markdown(f"Using column {student_col} for Student names.")
    st.markdown(f"Using column {school_col} for School names.")
    st.success(f" Loaded {len(participants)} participants.")

    # Generate button
    if st.button("Generate Certificates"):
        # Store user details in Supabase database
        data = {
            "name": user_name.strip(),
            "school_name": user_school_name.strip(),
            "school_number": user_school_number.strip(),
            "contact_number": user_contact_number.strip(),
            "ic_number": user_ic_number.strip(), 
            "timestamp": current_time,
            # MUST exist in Generations table!
            
        }
        
        try:
            response = supabase.table("Generations").insert(data).execute()
            
            # --- FIX APPLIED HERE ---
            if response.data:
                st.success("User details stored successfully.")
            elif response.error:
                 # Postgrest API Error
                st.error(f"Failed to store data in database. Supabase Error: {response.error['message']}")
            else:
                # Catch case where data is empty and no explicit error is returned
                 st.error("Failed to store data in database. Response data was empty.")
            # --- END FIX ---

        except Exception as e:
            st.error(f"Database insertion failed due to an unexpected error: {e}")
            # Ensure the process stops if database logging is critical
            # st.stop() # Uncomment this if database logging is mandatory for generation

        
        zip_buf = io.BytesIO()
        success, fail = 0, 0

        with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
            status_placeholder = st.empty() # Placeholder for status updates
            
            for idx, row in participants.iterrows():
                status_placeholder.info(f"Processing certificate {idx + 1} of {len(participants)}...")
                
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

                    # Read base template
                    base_reader = PdfReader(template_path)
                    base_page = base_reader.pages[0]
                    media_box = base_page.mediabox
                    w = float(media_box.width)
                    h = float(media_box.height)

                    # Create overlay
                    overlay_packet = io.BytesIO()
                    c = canvas.Canvas(overlay_packet, pagesize=(w, h))
                    c.setFillColorRGB(*hex_to_rgb(student_color))  
                    c.setFont(student_font, student_font_size)     
                    # The text is centered on the x-coordinate
                    c.drawCentredString(student_x, student_y, student)

                    if school:
                        c.setFillColorRGB(*hex_to_rgb(school_color))  
                        c.setFont(school_font, school_font_size)      
                        c.drawCentredString(school_x, school_y, school)

                    c.save()
                    overlay_packet.seek(0)

                    # Merge base + overlay correctly
                    overlay_reader = PdfReader(overlay_packet)
                    merged_page = PageObject.create_blank_page(
                        width=w, height=h
                    )
                    merged_page.merge_page(base_page)
                    merged_page.merge_page(overlay_reader.pages[0])

                    # Write output
                    out_buf = io.BytesIO()
                    writer = PdfWriter()
                    writer.add_page(merged_page)
                    writer.write(out_buf)
                    writer.close()
                    out_buf.seek(0)

                    # Safe filename
                    safe_name = re.sub(r'[<>:"/\\|?*]', "_", student.replace(" ", "_"))
                    filename = f"{idx+1:03d}_{safe_name}_certificate.pdf"

                    zipf.writestr(filename, out_buf.getvalue())
                    success += 1

                except Exception as e:
                    fail += 1
                    status_placeholder.error(f"Error creating certificate for {student or 'Unknown'}: {e}")
                    # Use a small sleep here if errors are flooding the UI, but generally not needed

        # Finalize and download
        status_placeholder.empty() # Clear the status update message
        zip_buf.seek(0)
        st.success(f"Generation Completed: **{success}** successful, **{fail}** failed.")

        st.download_button(
            "Download All Certificates (.zip)",
            data=zip_buf.getvalue(),
            file_name="certificates.zip",
            mime="application/zip"
        )

elif not inputs_valid:
    st.warning("Please fill in all required user details (Name, School Name, School Number, Contact Number) to proceed.")
else:
    st.info("Please upload the Excel file and fill in the required user details to begin.")
