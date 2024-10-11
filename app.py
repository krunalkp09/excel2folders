import streamlit as st
import pandas as pd
from fpdf import FPDF
from io import BytesIO

class PDF(FPDF):
    def header(self):
        # Add a large bold header with the sheet name
        self.set_font("Arial", "B", 16)
        self.cell(0, 10, f"{self.sheet_name}", 0, 1, "C")
        self.ln(5)

    def footer(self):
        # Go to 1.5 cm from bottom
        self.set_y(-15)
        # Select Arial italic 8
        self.set_font("Arial", "I", 8)
        
        # Display unique subdivisions for the current page
        if hasattr(self, "current_subdivisions") and self.current_subdivisions:
            subdivisions_text = "Subdivisions: " + ", ".join(self.current_subdivisions)
            self.cell(0, 10, subdivisions_text, 0, 0, "C")
        else:
            self.cell(0, 10, "Subdivision: Not Available", 0, 0, "C")

# Step 1: File Upload
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # Step 2: Load Excel and Select Sheet
    excel_data = pd.ExcelFile(uploaded_file)
    sheet_names = excel_data.sheet_names
    sheet = st.selectbox("Select a sheet to process", sheet_names)

    # Step 3: Load Data from Selected Sheet
    if sheet:
        df = excel_data.parse(sheet)
        st.write("Preview of the data:")
        st.write(df.head())

        # Extract specific columns if they exist
        required_columns = ["Last_Name", "First_Name", "Address", "City"]
        if all(col in df.columns for col in required_columns):
            # Step 4: PDF Generation
            if st.button("Generate PDF"):
                pdf = PDF()
                pdf.sheet_name = sheet
                
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                
                # Collect unique subdivisions for each page
                current_page_subdivisions = set()
                records_per_page = 12  # Set limit to 10 records per page
                record_count = 0  # Track the number of records on the current page
                
                # Set box dimensions to utilize page height
                name_address_width = 85  # Narrower width for name/address
                comment_width = 110  # Wider width for comments
                box_height = 20  # Reduced height for each entry to remove spacing
                
                # Add each row to the PDF without any extra vertical space
                for index, row in df.iterrows():
                    # Track subdivisions for the current page
                    if "Subdivision_Name" in df.columns:
                        current_page_subdivisions.add(row["Subdivision_Name"])
                    
                    # Draw boxes for the name/address and comments
                    pdf.set_xy(10, pdf.get_y())
                    pdf.cell(name_address_width, box_height, border=1)  # Name/Address box
                    pdf.cell(comment_width, box_height, border=1)  # Comments box
                    
                    # Inside the name/address box
                    pdf.set_xy(10, pdf.get_y())
                    full_name = f"{row['Last_Name']}, {row['First_Name']}"
                    pdf.cell(name_address_width, 7, txt=full_name, border=0)
                    
                    pdf.set_xy(10, pdf.get_y() + 7)
                    address_line = row["Address"]
                    pdf.cell(name_address_width, 7, txt=address_line, border=0)
                    
                    pdf.set_xy(10, pdf.get_y() + 7)
                    city_line = row["City"]
                    pdf.cell(name_address_width, 6, txt=city_line, border=0)

                    # Inside the comments box with "Comments:" label
                    pdf.set_xy(10 + name_address_width, pdf.get_y() - 14)
                    pdf.cell(comment_width, 7, txt="Comments:", border=0)
                    
                    # Add three end-to-end lines for comments
                    line_y = pdf.get_y() + 7
                    pdf.set_xy(10 + name_address_width, line_y)
                    for _ in range(2):  # Three lines for comments
                        pdf.cell(comment_width, 6, txt="_________________________________________", border=0)
                        line_y += 6
                        pdf.set_xy(10 + name_address_width, line_y)

                    # Adjust y-position for next entry
                    pdf.set_y(pdf.get_y() - box_height + 22)
                    record_count += 1

                    # If record count reaches the limit, add a page and reset count
                    if record_count >= records_per_page:
                        pdf.current_subdivisions = list(current_page_subdivisions)
                        pdf.add_page()
                        current_page_subdivisions.clear()
                        record_count = 0

                # Final page subdivisions
                pdf.current_subdivisions = list(current_page_subdivisions)
                
                # Output PDF content as string and convert to BytesIO for download
                pdf_output = pdf.output(dest='S').encode('latin1')  # Convert to bytes
                pdf_buffer = BytesIO(pdf_output)
                
                st.download_button("Download PDF", data=pdf_buffer, file_name="formatted_report_no_spacing.pdf", mime="application/pdf")
        else:
            st.warning(f"Columns {required_columns} are required in the sheet.")
