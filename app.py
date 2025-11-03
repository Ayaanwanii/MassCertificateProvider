import streamlit as st
import pandas as pd
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io, re, zipfile, os, datetime
from supabase import create_client, Client

# --- FONT SETUP ---
# NOTE: Ensure 'Bliss Extra Bold.ttf' is in the root directory for deployment.
try:
    # Register font (ensure this font file exists in your working folder)
    pdfmetrics.registerFont(TTFont('BlissExtraBold', './Bliss Extra Bold.ttf'))
except Exception as e:
    st.warning(f"Warning: Could not register custom font. Ensure 'Bliss Extra Bold.ttf' is in the root directory. Using default fonts for now. Error: {e}")

# --- SUPABASE SETUP ---
# Requires Streamlit secrets: st.secrets["supabase"]["url"] and st.secrets["supabase"]["key"]
SUPABASE_URL = st.secrets["supabase"]["url"]
SUPABASE_KEY = st.secrets["supabase"]["key"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

def init_db():
    pass 

init_db()

# --- INITIALIZE SESSION STATE ---
# This dictionary stores data across Streamlit reruns
if 'details_submitted' not in st.session_state:
    st.session_state.details_submitted = False
if 'user_data' not in st.session_state:
    st.session_state.user_data = {}


# --- STREAMLIT PAGE SETUP ---
st.set_page_config(page_title="Certificate Generator", page_icon="🎓", layout="wide")
st.title("Automated Certificate Generator")


# Fixed certificate template path
template_path = './certificate_template.pdf'
if not os.path.exists(template_path):
    st.error("Certificate template not found. Please ensure 'certificate_template.pdf' is in the working directory.")
    st.stop()

# --- UTILITY FUNCTION ---
def hex_to_rgb(hex_color):
    """Convert hex color to RGB tuple (0-1 range for ReportLab)."""
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) / 255 for i in (0, 2, 4))

# --- CERTIFICATE STYLES (HARDCODED) ---
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

# --- USER INPUT FIELDS INSIDE A FORM ---
st.markdown("---")
st.markdown("1. Submit Your Details to Download Certificates")

# Use a form to require a submit button press
with st.form("user_details_form"):
    # Pre-populate inputs from session state if already submitted
    user_name = st.text_input("Name", st.session_state.user_data.get('name', ''))
    user_school_name = st.text_input("School Name", st.session_state.user_data.get('school_name', ''))
    user_school_number = st.text_input("School Number", st.session_state.user_data.get('school_number', ''))
    user_contact_number = st.text_input("Contact Number", st.session_state.user_data.get('contact_number', ''))
    user_ic_number = st.text_input("IC Number", st.session_state.user_data.get('ic_number', '')) 

    # Validate all required fields
    inputs_valid = all([
        user_name.strip(), 
        user_school_name.strip(), 
        user_school_number.strip(), 
        user_contact_number.strip(), 
        user_ic_number.strip()
    ])

    submitted = st.form_submit_button("Submit Details")

if submitted:
    if inputs_valid:
        
        # Prepare data for Supabase
        user_data = {
            "name": user_name.strip(),
            "school_name": user_school_name.strip(),
            "school_number": user_school_number.strip(),
            "contact_number": user_contact_number.strip(),
            "ic_number": user_ic_number.strip(),
            "timestamp": datetime.datetime.now(datetime.timezone.utc).isoformat()
        }
        
        # --- NEW DATABASE INSERTION BLOCK ---
        with st.spinner("Saving details..."):
            try:
                response = supabase.table("Generations").insert(user_data).execute()
                
                if response.data:
                    # Save to session state only after successful database insertion
                    st.session_state.details_submitted = True
                    st.session_state.user_data = user_data 
                    st.success("User details submitted and stored successfully! You can now upload your Excel/CSV file and generate certificates.")
                    st.balloons()
                elif response.error:
                    st.session_state.details_submitted = False
                    st.error(f"Failed to store data in database. Supabase Error: {response.error['message']}")
                else:
                    st.session_state.details_submitted = False
                    st.error("Failed to store data in database. Response data was empty.")

            except Exception as e:
                st.session_state.details_submitted = False
                st.error(f"Database connection or insertion failed: {e}")
        # --- END NEW DATABASE INSERTION BLOCK ---

    else:
        st.session_state.details_submitted = False
        st.error("Please fill in ALL required user details before submitting.")

st.markdown("---")

# --- MAIN LOGIC BLOCK (Checks for submitted state) ---
if st.session_state.details_submitted:
    
    st.markdown("2. Generate Certificates")

    # File upload - MODIFIED TO ACCEPT CSV AND XLSX
    excel_file = st.file_uploader("Upload Participant List (Excel/CSV)", type=["xlsx", "csv"])

    # --- START OF FILE-DEPENDENT CODE BLOCK (Fixes the NameError) ---
    if excel_file:
        
        try:
            # Conditional reading based on file type
            file_name = excel_file.name.lower()
            
            if file_name.endswith('.csv'):
                participants = pd.read_csv(excel_file, header=0)
            elif file_name.endswith(('.xlsx', '.xls')):
                participants = pd.read_excel(excel_file, header=0)
            else:
                st.error("Unsupported file format detected. Please upload an XLSX or CSV file.")
                st.stop()
            
        except Exception as e:
            st.error(f"Error reading file: {e}. Please check your file format and data integrity.")
            st.stop() # Stop execution if file reading fails

        # --- ALL CODE THAT USES 'participants' MUST BE INSIDE THIS BLOCK ---
        
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

        st.info(f"Ready to process {len(participants)} participants. Using `{student_col}` for names.")

        # --- GENERATE BUTTON (APPEARS ONLY WHEN FILE IS UPLOADED AND PROCESSED) ---
        if st.button("Generate Certificates and Download Zip"):
            
            # Status update immediately after button click
            with st.spinner("Starting certificate generation..."): 
                st.caption("Initiating PDF merging and compression...")
            
            # --- CERTIFICATE GENERATION ---
            zip_buf = io.BytesIO()
            success, fail = 0, 0

            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zipf:
                status_placeholder = st.empty() 
                
                for idx, row in participants.iterrows():
                    status_placeholder.info(f"Processing certificate {idx + 1} of {len(participants)} for: {row[student_col]}...")
                    
                    try:
                        raw_student = row[student_col]
                        if pd.isna(raw_student) or str(raw_student).strip() == "":
                            fail += 1
                            continue
                        student = str(raw_student).strip()

                        school = ""
                        if school_col is not None:
                            raw_school = row[school_col]
                            if not pd.isna(raw_school) and str(raw_school).strip() != "":
                                school = str(raw_school).strip()

                        # PDF Processing
                        base_reader = PdfReader(template_path)
                        base_page = base_reader.pages[0]
                        media_box = base_page.mediabox
                        w = float(media_box.width)
                        h = float(media_box.height)

                        # Create overlay canvas
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

                        # Merge PDF pages
                        overlay_reader = PdfReader(overlay_packet)
                        merged_page = PageObject.create_blank_page(width=w, height=h)
                        merged_page.merge_page(base_page)
                        merged_page.merge_page(overlay_reader.pages[0])

                        # Write output PDF
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
            status_placeholder.empty() 
            zip_buf.seek(0)
            
            # --- FINAL SUCCESS MESSAGE AND DOWNLOAD BUTTON ---
            st.balloons()
            st.success(f"Generation Completed! {success} successful, {fail} failed.")

            st.download_button(
                "Download All Certificates (.zip)",
                data=zip_buf.getvalue(),
                file_name="certificates.zip",
                mime="application/zip"
            )
            # --- END FINAL ---
    else:
        # Message if details are submitted but file is missing
        st.info("Please upload the Participant List (Excel/CSV) to proceed with generation.")


# --- VALIDATION MESSAGES ---
elif not st.session_state.details_submitted:
    st.warning("Please submit your user details using the **Submit Details & Save to Database** button in Section 1.")
