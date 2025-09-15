import streamlit as st
import pandas as pd
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Border, Side, Alignment, PatternFill
import io
from datetime import datetime

def copy_cell_style(source_cell, target_cell):
    """Copy style from source cell to target cell"""
    try:
        if source_cell.font:
            target_cell.font = Font(
                name=source_cell.font.name,
                size=source_cell.font.size,
                bold=source_cell.font.bold,
                italic=source_cell.font.italic,
                color=source_cell.font.color
            )
        if source_cell.border:
            target_cell.border = Border(
                left=source_cell.border.left,
                right=source_cell.border.right,
                top=source_cell.border.top,
                bottom=source_cell.border.bottom
            )
        if source_cell.fill:
            target_cell.fill = PatternFill(
                fill_type=source_cell.fill.fill_type,
                start_color=source_cell.fill.start_color,
                end_color=source_cell.fill.end_color
            )
        if source_cell.alignment:
            target_cell.alignment = Alignment(
                horizontal=source_cell.alignment.horizontal,
                vertical=source_cell.alignment.vertical,
                wrap_text=source_cell.alignment.wrap_text
            )
    except:
        pass  # Skip if there's any error with styling

def format_date_value(date_val):
    """Format date value to string"""
    if pd.isna(date_val) or date_val is None:
        return ""
    
    if isinstance(date_val, str):
        return date_val
    
    # If it's a datetime object
    if hasattr(date_val, 'strftime'):
        return date_val.strftime('%d-%m-%Y')
    
    return str(date_val)

def generate_forms(data_df, template_file):
    """Generate multiple forms from data"""
    
    try:
        # Load template
        template_wb = load_workbook(template_file)
        template_ws = template_wb.active
        
        # Create output workbook
        output_wb = Workbook()
        output_ws = output_wb.active
        output_ws.title = "Generated Forms"
        
        # Find template boundaries
        max_row = 0
        max_col = 0
        for row in template_ws.iter_rows():
            for cell in row:
                if cell.value is not None:
                    max_row = max(max_row, cell.row)
                    max_col = max(max_col, cell.column)
        
        st.write(f"Template size: {max_row} rows x {max_col} columns")
        
        current_output_row = 1
        form_spacing = 5  # Space between forms
        
        # Process each data record
        for record_idx, data_row in data_df.iterrows():
            st.write(f"Processing record {record_idx + 1}/{len(data_df)}")
            st.write(f"Available columns: {list(data_row.index)}")
            
            # Copy template structure first
            for template_row in range(1, max_row + 1):
                for template_col in range(1, max_col + 1):
                    source_cell = template_ws.cell(row=template_row, column=template_col)
                    target_cell = output_ws.cell(row=current_output_row + template_row - 1, column=template_col)
                    
                    # Copy value and style
                    target_cell.value = source_cell.value
                    copy_cell_style(source_cell, target_cell)
            
            # Now fill in the data
            data_filled_count = 0
            
            for template_row in range(1, max_row + 1):
                # Get all cells in this row to analyze
                row_cells = []
                for col in range(1, max_col + 1):
                    cell = template_ws.cell(row=template_row, column=col)
                    row_cells.append((col, cell.value))
                
                # Find label (usually in first few columns)
                label = None
                label_col = None
                for col, value in row_cells:
                    if value and isinstance(value, str) and len(value.strip()) > 0:
                        # Skip cells that are just colons or dots
                        if value.strip() not in [':', ':.', ':..', ':...']:
                            label = value.strip().lower()
                            label_col = col
                            break
                
                if not label:
                    continue
                
                st.write(f"Row {template_row}: Found label '{label}' in column {label_col}")
                
                # Find colon/dots column (where to put data)
                data_col = None
                for col, value in row_cells:
                    if value and ':' in str(value):
                        data_col = col
                        st.write(f"Found data column {data_col} with value: '{value}'")
                        break
                
                if not data_col:
                    st.write(f"No data column found for row {template_row}")
                    continue
                
                # Fill data based on label
                target_cell = output_ws.cell(row=current_output_row + template_row - 1, column=data_col)
                original_value = str(target_cell.value) if target_cell.value else ""
                
                filled = False
                
                # Map data
                if 'nama bayi' in label or 'nama balita' in label:
                    for col_name in ['Nama Bayi/Balita', 'NAMA ANAK', 'NAMA BAYI']:
                        if col_name in data_row.index and pd.notna(data_row[col_name]):
                            target_cell.value = f": {data_row[col_name]}"
                            filled = True
                            st.write(f"✅ Filled Nama: {data_row[col_name]}")
                            break
                
                elif 'nik' in label:
                    if 'NIK' in data_row.index and pd.notna(data_row['NIK']):
                        nik_str = str(data_row['NIK']).replace('.0', '')
                        target_cell.value = f": {nik_str}"
                        target_cell.number_format = '@'  # Text format
                        filled = True
                        st.write(f"✅ Filled NIK: {nik_str}")
                
                elif 'tanggal' in label and 'lahir' in label:
                    for col_name in ['TANGGAL LAHIR', 'TGL LAHIR']:
                        if col_name in data_row.index and pd.notna(data_row[col_name]):
                            date_str = format_date_value(data_row[col_name])
                            if date_str:
                                target_cell.value = f": {date_str}"
                                filled = True
                                st.write(f"✅ Filled Tanggal: {date_str}")
                            break
                
                elif 'berat badan' in label:
                    if 'BB' in data_row.index and pd.notna(data_row['BB']):
                        target_cell.value = f": {data_row['BB']}"
                        filled = True
                        st.write(f"✅ Filled BB: {data_row['BB']}")
                
                elif 'panjang badan' in label or 'tinggi badan' in label:
                    if 'TB' in data_row.index and pd.notna(data_row['TB']):
                        target_cell.value = f": {data_row['TB']}"
                        filled = True
                        st.write(f"✅ Filled TB: {data_row['TB']}")
                
                elif 'nama ayah' in label and 'ibu' in label:
                    # Combined field, put Ayah first
                    if 'AYAH' in data_row.index and pd.notna(data_row['AYAH']):
                        target_cell.value = f": {data_row['AYAH']}"
                        filled = True
                        st.write(f"✅ Filled Ayah: {data_row['AYAH']}")
                        
                        # Try to fill Ibu in next row
                        next_row = template_row + 1
                        if next_row <= max_row:
                            # Find data column in next row
                            for col in range(1, max_col + 1):
                                next_cell = template_ws.cell(row=next_row, column=col)
                                if next_cell.value and ':' in str(next_cell.value):
                                    next_target = output_ws.cell(row=current_output_row + next_row - 1, column=col)
                                    if 'IBU' in data_row.index and pd.notna(data_row['IBU']):
                                        next_target.value = f": {data_row['IBU']}"
                                        st.write(f"✅ Filled Ibu: {data_row['IBU']}")
                                    break
                
                elif 'alamat' in label:
                    if 'ALAMAT' in data_row.index and pd.notna(data_row['ALAMAT']):
                        target_cell.value = f": {data_row['ALAMAT']}"
                        filled = True
                        st.write(f"✅ Filled Alamat: {data_row['ALAMAT']}")
                
                elif 'no' in label and 'hp' in label:
                    for col_name in ['No. Hp', 'NO HP', 'NOHP']:
                        if col_name in data_row.index and pd.notna(data_row[col_name]):
                            target_cell.value = f": {data_row[col_name]}"
                            filled = True
                            st.write(f"✅ Filled HP: {data_row[col_name]}")
                            break
                
                if filled:
                    data_filled_count += 1
                else:
                    st.write(f"❌ No data found for: '{label}'")
            
            # Handle gender selection
            gender_value = None
            for col_name in ['JENIS', 'JENIS KELAMIN', 'L/P', 'GENDER']:
                if col_name in data_row.index and pd.notna(data_row[col_name]):
                    gender_value = str(data_row[col_name]).upper()
                    break
            
            if gender_value:
                st.write(f"Processing gender: {gender_value}")
                for template_row in range(1, max_row + 1):
                    for template_col in range(1, max_col + 1):
                        source_cell = template_ws.cell(row=template_row, column=template_col)
                        if (source_cell.value and 
                            'laki' in str(source_cell.value).lower() and 
                            'perempuan' in str(source_cell.value).lower()):
                            
                            target_cell = output_ws.cell(row=current_output_row + template_row - 1, column=template_col)
                            
                            if gender_value in ['L', 'LAKI', 'LAKI-LAKI']:
                                target_cell.value = "Laki-laki"
                                st.write("✅ Set gender: Laki-laki")
                            elif gender_value in ['P', 'PEREMPUAN']:
                                target_cell.value = "Perempuan"
                                st.write("✅ Set gender: Perempuan")
                            break
            
            st.write(f"Form {record_idx + 1} completed. Data fields filled: {data_filled_count}")
            st.write("---")
            
            # Move to next form position
            current_output_row += max_row + form_spacing
        
        # Save to BytesIO
        output_file = io.BytesIO()
        output_wb.save(output_file)
        output_file.seek(0)
        
        return output_file
        
    except Exception as e:
        st.error(f"Error in generate_forms: {str(e)}")
        raise e

# Streamlit App
def main():
    st.set_page_config(
        page_title="Excel Form Generator", 
        page_icon="📋",
        layout="wide"
    )
    
    st.title("📋 Excel Form Generator")
    st.markdown("Upload your data file and form template to generate multiple forms automatically!")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("1. Upload Data File")
        data_file = st.file_uploader(
            "Choose your data Excel file", 
            type=['xlsx', 'xls'],
            help="File containing the records to be processed"
        )
        
        if data_file:
            try:
                data_df = pd.read_excel(data_file)
                st.success(f"✅ Data loaded: {len(data_df)} records found")
                
                # Show preview
                st.subheader("Data Preview")
                st.dataframe(data_df.head(3), use_container_width=True)
                
                # Show column info
                st.subheader("Available Columns")
                cols_display = []
                for i, col in enumerate(data_df.columns):
                    cols_display.append(f"{i+1}. '{col}'")
                st.text("\n".join(cols_display))
                
            except Exception as e:
                st.error(f"Error reading data file: {str(e)}")
                data_df = None
    
    with col2:
        st.subheader("2. Upload Form Template")
        template_file = st.file_uploader(
            "Choose your form template Excel file", 
            type=['xlsx', 'xls'],
            help="Template form that will be filled with data"
        )
        
        if template_file:
            try:
                # Just check if it loads
                test_wb = load_workbook(template_file)
                st.success("✅ Template loaded successfully")
                
                # Show basic info
                ws = test_wb.active
                max_row = ws.max_row
                max_col = ws.max_column
                st.write(f"Template size: {max_row} rows × {max_col} columns")
                
            except Exception as e:
                st.error(f"Error reading template file: {str(e)}")
    
    # Generate forms section
    if data_file and template_file:
        st.markdown("---")
        st.subheader("3. Generate Forms")
        
        # Test with first record only
        if st.checkbox("Test with first record only (recommended)"):
            test_df = data_df.head(1)
        else:
            test_df = data_df
        
        if st.button("🚀 Generate Forms", type="primary", use_container_width=True):
            try:
                with st.spinner("Generating forms... Please wait."):
                    output_file = generate_forms(test_df, template_file)
                
                st.success(f"✅ Successfully generated {len(test_df)} forms!")
                
                # Download button
                st.download_button(
                    label="📥 Download Generated Forms",
                    data=output_file.getvalue(),
                    file_name=f"Generated_Forms_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
            except Exception as e:
                st.error(f"Error generating forms: {str(e)}")
                st.exception(e)
    
    # Instructions
    with st.expander("📖 Instructions"):
        st.markdown("""
        ### Quick Start:
        1. Upload your **data Excel file** (contains records)
        2. Upload your **form template Excel file** (the blank form)
        3. Check "Test with first record only" for testing
        4. Click "Generate Forms"
        5. Check the debug output to see what's happening
        6. Download the result
        
        ### Expected Data Columns:
        - `Nama Bayi/Balita` or `NAMA ANAK`
        - `NIK`
        - `TANGGAL LAHIR`
        - `BB` (Berat Badan)
        - `TB` (Tinggi Badan)
        - `AYAH`, `IBU`
        - `ALAMAT`
        - `No. Hp`
        - `JENIS` (L/P for gender)
        
        ### Template Requirements:
        - Labels in left columns (Nama Bayi/Balita, NIK, etc.)
        - Colon with dots (`:....`) in cells where data should go
        - `(Laki-laki/Perempuan)` text for gender selection
        """)

if __name__ == "__main__":
    main