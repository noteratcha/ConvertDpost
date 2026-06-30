import os
import re
import glob
import pandas as pd
from pypdf import PdfReader

__version__ = "2026.0630.1453"

# Mapping of Thai digits to Arabic digits
THAI_TO_ARABIC = str.maketrans('๐๑๒๓๔๕๖๗๘๙', '0123456789')

def clean_thai_digits(text):
    if not text:
        return ""
    return text.translate(THAI_TO_ARABIC)

def extract_tel(text):
    """Extract telephone number from text and convert to Arabic digits."""
    if not text:
        return ""
    # Find patterns like โทร. 08-xxxx-xxxx or โทร. ๐๘ xxxx xxxx
    match = re.search(r'โทร\s*\.?\s*([๐-๙0-9\-\s\.,]+)', text)
    if match:
        tel = match.group(1).strip()
        # Clean digits and remove non-numeric chars except hyphen
        tel_clean = clean_thai_digits(tel)
        tel_digits = re.sub(r'[^0-9]', '', tel_clean)
        return tel_digits
    return ""

def parse_address_components(address_text):
    """
    Parse Tambon, Amphur, Province and Zipcode from address text.
    Standard patterns:
    ตำบล/แขวง ... -> ต. ...
    อำเภอ/เขต ... -> อ. ...
    จังหวัด ... -> จ. ...
    """
    # Normalize spaces
    address_text = re.sub(r'\s+', ' ', address_text)
    
    # Extract zipcode (5 digits)
    zip_match = re.search(r'(\b[0-9]{5}\b)', address_text)
    zipcode = zip_match.group(1) if zip_match else ""
    
    # Extract Amphur
    amphur_match = re.search(r'(?:อำเภอ/เขต|อำเภอ|อ\.)\s*([^\sจ\.]+)', address_text)
    amphur = amphur_match.group(1).strip() if amphur_match else ""
    
    # Extract Province
    province_match = re.search(r'(?:จังหวัด|จ\.)\s*([^\s\d]+)', address_text)
    province = province_match.group(1).strip() if province_match else ""
    
    # Fallback for Province: if not found, look for the word immediately preceding the zipcode
    if not province and zipcode:
        prov_match = re.search(r'([ก-ฮ]{2,})\s+' + zipcode, address_text)
        if prov_match:
            province = prov_match.group(1).strip()
            
    return amphur, province, zipcode

def parse_receiver_label(text):
    """
    Parse the receiver label block:
    เรียน <Name>
    <Address lines>
    <5-digit zipcode>
    """
    # Normalize lines
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    receiver_name = ""
    address_lines = []
    zipcode = ""
    
    # Look for the index of the line starting with "เรียน"
    start_idx = -1
    for i, line in enumerate(lines):
        if line.startswith("เรียน"):
            start_idx = i
            break
            
    if start_idx == -1:
        return None
        
    receiver_name = re.sub(r'^เรียน\s*', '', lines[start_idx]).strip()
    
    # Collect subsequent lines until a zipcode is found
    zip_pattern = re.compile(r'^[๐-๙0-9]{5}$')
    end_idx = -1
    
    for i in range(start_idx + 1, len(lines)):
        line_clean = clean_thai_digits(lines[i])
        if zip_pattern.match(line_clean):
            zipcode = line_clean
            end_idx = i
            break
            
    if end_idx == -1:
        # If no strict 5-digit line, check if the last line contains a zipcode
        for i in range(start_idx + 1, len(lines)):
            line_clean = clean_thai_digits(lines[i])
            zip_match = re.search(r'([0-9]{5})', line_clean)
            if zip_match:
                zipcode = zip_match.group(1)
                end_idx = i
                break
                
    if end_idx == -1:
        return None # Could not find valid zipcode boundary
        
    # Address lines are between start_idx+1 and end_idx (inclusive if zip was at the end of a line)
    raw_address_lines = lines[start_idx + 1 : end_idx]
    # If the end line had text before zip, append the text part
    end_line_clean = lines[end_idx]
    if len(clean_thai_digits(end_line_clean)) > 5:
        # Remove zipcode from the line
        text_before_zip = re.sub(r'[๐-๙0-9]{5}\s*$', '', end_line_clean).strip()
        if text_before_zip:
            raw_address_lines.append(text_before_zip)
            
    # Clean address lines (e.g. replace ตำบล/แขวง with ต. , อำเภอ/เขต with อ.)
    cleaned_address_lines = []
    amphur = ""
    province = ""
    
    for line in raw_address_lines:
        line_conv = clean_thai_digits(line)
        
        # Check for tambon
        t_match = re.search(r'^(?:ตำบล/แขวง|ตำบล|ต\.)\s*(.+)$', line_conv)
        if t_match:
            cleaned_address_lines.append(f"ต.{t_match.group(1).strip()}")
            continue
            
        # Check for amphur
        a_match = re.search(r'^(?:อำเภอ/เขต|อำเภอ|อ\.)\s*(.+)$', line_conv)
        if a_match:
            amphur = a_match.group(1).strip()
            # Do NOT append to cleaned_address_lines to cut from receiver_address
            continue
            
        # Check for province
        p_match = re.search(r'^(?:จังหวัด|จ\.)\s*(.+)$', line_conv)
        if p_match:
            province = p_match.group(1).strip()
            # Do NOT append to cleaned_address_lines to cut from receiver_address
            continue
            
        cleaned_address_lines.append(line)
        
    # Reconstruct address (excluding zipcode to cut it from receiver_address)
    receiver_address = " ".join(cleaned_address_lines)
    
    # Try to extract amphur and province if not found yet
    if not amphur or not province:
        raw_joined = " ".join(raw_address_lines) + " " + zipcode
        ext_amphur, ext_province, _ = parse_address_components(raw_joined)
        if not amphur:
            amphur = ext_amphur
        if not province:
            province = ext_province
            
    # Post-processing cleanup: Remove amphur, province, and zipcode from receiver_address if present
    if amphur:
        receiver_address = re.sub(r'(?:อำเภอ/เขต|อำเภอ|อ\.)\s*' + re.escape(amphur), '', receiver_address)
        receiver_address = re.sub(r'\b' + re.escape(amphur) + r'\b', '', receiver_address)
    if province:
        receiver_address = re.sub(r'(?:จังหวัด|จ\.)\s*' + re.escape(province), '', receiver_address)
        receiver_address = re.sub(r'\b' + re.escape(province) + r'\b', '', receiver_address)
    if zipcode:
        receiver_address = receiver_address.replace(zipcode, "")
        
    # Normalize spaces and strip
    receiver_address = re.sub(r'\s+', ' ', receiver_address).strip()
            
    return {
        'RECEIVER': receiver_name,
        'RECEIVER ADDRESS': receiver_address,
        'RECEIVER AMPHUR': amphur,
        'RECEIVER PROVINCE': province,
        'RECEIVER ZIPCODE': zipcode
    }

def parse_shipper_label(text):
    """
    Parse raw shipper info from mailing label block:
    ฝ่ายรังวัด สำนักงานที่ดินจังหวัดนครพนม สาขาเรณูนคร
    อำเภอเรณูนคร นครพนม ๔๘๑๗๐
    ที่ นพ๐๐๒๐.๐๕/๘๘๙
    """
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    
    shipper_name = ""
    shipper_address = ""
    ref_no = ""
    
    # Find the shipper line
    shipper_idx = -1
    for i, line in enumerate(lines):
        if "ฝ่ายรังวัด สำนักงานที่ดิน" in line or "สำนักงานที่ดิน" in line:
            shipper_idx = i
            break
            
    if shipper_idx != -1:
        shipper_name = lines[shipper_idx]
        
        # Collect subsequent lines as address parts until we hit reference number or other block
        addr_parts = []
        for j in range(shipper_idx + 1, len(lines)):
            line = lines[j]
            # Stop if we hit reference number, receiver, or another major section
            if line.startswith("ที่") or line.startswith("เรียน") or line.startswith("วันที่") or "ผู้ขอรังวัด" in line or "นายช่างรังวัด" in line or "ฝ่ายรังวัด" in line or "ท.ด." in line:
                break
            addr_parts.append(line)
            
        if addr_parts:
            raw_addr = " ".join(addr_parts)
            # Clean address
            addr_conv = clean_thai_digits(raw_addr)
            # Ensure correct formatting (e.g. อ.เรณูนคร จ.นครพนม)
            addr_conv = re.sub(r'อำเภอ/เขต|อำเภอ|อ\.', 'อ.', addr_conv)
            addr_conv = re.sub(r'จังหวัด|จ\.', 'จ.', addr_conv)
            shipper_address = addr_conv
                
        # Look for reference number "ที่ ..."
        for line in lines[shipper_idx:]:
            if line.startswith("ที่"):
                ref_no = clean_thai_digits(line.replace("ที่", "").strip())
                break
                
    return {
        'SHIPPER NAME': shipper_name,
        'SHIPPER ADDRESS': shipper_address,
        'REF NO': ref_no
    }

def process_pdf(pdf_path):
    print(f"กำลังประมวลผลไฟล์: {os.path.basename(pdf_path)}...")
    reader = PdfReader(pdf_path)
    records = []
    
    # 1. Scan all pages to extract the best/most complete shipper info
    best_shipper_name = ""
    best_shipper_address = ""
    best_shipper_tel = ""
    
    for page in reader.pages:
        text = page.extract_text() or ""
        shipper_info = parse_shipper_label(text)
        if shipper_info:
            name = shipper_info.get('SHIPPER NAME', '')
            addr = shipper_info.get('SHIPPER ADDRESS', '')
            if len(name) > len(best_shipper_name):
                best_shipper_name = name
            if len(addr) > len(best_shipper_address):
                best_shipper_address = addr
                
        tel = extract_tel(text)
        if len(tel) > len(best_shipper_tel):
            best_shipper_tel = tel
            
    # 2. Post-process the shipper info based on the rules:
    # Rule 2: "ถ้าข้อมูลที่คอลัมน์ G (SHIPPER ADDRESS) ว่าง ให้ไปเอาข้อมูลวรรคสุดท้ายของ F (SHIPPER NAME) มาใส่ และลบข้อความนั้นออกจากคอลัมน์ F"
    if not best_shipper_address and best_shipper_name:
        parts = [p.strip() for p in best_shipper_name.split() if p.strip()]
        if len(parts) > 1:
            last_part = parts[-1]
            best_shipper_address = last_part
            best_shipper_name = " ".join(parts[:-1])
            
    # Extract components from shipper address
    shipper_amphur = ""
    shipper_province = ""
    shipper_zipcode = ""
    
    if best_shipper_address:
        shipper_amphur, shipper_province, shipper_zipcode = parse_address_components(best_shipper_address)
        
    # Fallback to parse Amphur and Province from SHIPPER NAME if still missing
    # E.g. "สำนักงานที่ดินจังหวัดนครพนม สาขาเรณูนคร"
    if not shipper_amphur and best_shipper_name:
        branch_match = re.search(r'สาขา\s*([ก-ฮ]+)', best_shipper_name)
        if branch_match:
            shipper_amphur = branch_match.group(1).strip()
            
    if not shipper_province and best_shipper_name:
        prov_match = re.search(r'สำนักงานที่ดินจังหวัด\s*([ก-ฮ]+)', best_shipper_name)
        if prov_match:
            shipper_province = prov_match.group(1).strip()
            
    # Rule 1: "ถ้าเอาข้อมูลบางส่วนมาระบุที่คอลัมน์ H,I,J,K (AMPHUR, PROVINCE, ZIPCODE, TEL) แล้วให้ลบข้อความนั้นออกจากคอลัมน์ G (SHIPPER ADDRESS)"
    if best_shipper_address:
        if shipper_amphur:
            best_shipper_address = re.sub(r'(?:อำเภอ/เขต|อำเภอ|อ\.)\s*' + re.escape(shipper_amphur), '', best_shipper_address)
            best_shipper_address = re.sub(r'\b' + re.escape(shipper_amphur) + r'\b', '', best_shipper_address)
        if shipper_province:
            best_shipper_address = re.sub(r'(?:จังหวัด|จ\.)\s*' + re.escape(shipper_province), '', best_shipper_address)
            best_shipper_address = re.sub(r'\b' + re.escape(shipper_province) + r'\b', '', best_shipper_address)
        if shipper_zipcode:
            best_shipper_address = best_shipper_address.replace(str(shipper_zipcode), "")
        if best_shipper_tel:
            best_shipper_address = best_shipper_address.replace(str(best_shipper_tel), "")
            
        # Clean up spaces
        best_shipper_address = re.sub(r'\s+', ' ', best_shipper_address).strip()
        
    # Prepare the structured shipper dict
    final_shipper_info = {
        'SHIPPER NAME': best_shipper_name,
        'SHIPPER ADDRESS': best_shipper_address,
        'SHIPPER AMPHUR': shipper_amphur,
        'SHIPPER PROVINCE': shipper_province,
        'SHIPPER ZIPCODE': shipper_zipcode,
        'SHIPPER TEL': best_shipper_tel
    }
    
    # 3. Process each page to associate receiver details with the best shipper info
    for i in range(len(reader.pages)):
        text = reader.pages[i].extract_text() or ""
        
        # Check if page contains receiver address block (mailing label)
        if "เรียน" in text and any(char.isdigit() or char in "๐๑๒๓๔๕๖๗๘๙" for char in text):
            # Parse receiver details
            receiver_info = parse_receiver_label(text)
            if receiver_info:
                # Find REF NO (Reference Number) for this specific letter
                # Scan current page
                page_shipper = parse_shipper_label(text)
                ref_no = page_shipper.get('REF NO', '')
                
                # Fallback to preceding page if REF NO is empty
                if not ref_no and i > 0:
                    prev_text = reader.pages[i - 1].extract_text() or ""
                    prev_shipper = parse_shipper_label(prev_text)
                    ref_no = prev_shipper.get('REF NO', '')
                    
                # Combine parsed information
                record = {
                    **final_shipper_info,
                    **receiver_info,
                    'REF NO': ref_no
                }
                records.append(record)
                
    return records

def records_to_dataframe(all_records):
    # Define exact columns matching the DPost template
    columns = [
        'NO', 'COMP ORDER ID', 'REF NO', 'BARCODE NO', 'PRODUCT IN BOX',
        'SHIPPER NAME', 'SHIPPER ADDRESS', 'SHIPPER AMPHUR', 'SHIPPER PROVINCE', 'SHIPPER ZIPCODE',
        'SHIPPER TEL', 'SHIPPER EMAIL', 'RECEIVER', 'RECEIVER ADDRESS', 'RECEIVER AMPHUR',
        'RECEIVER PROVINCE', 'RECEIVER ZIPCODE', 'RECEIVER TEL', 'RECEIVER EMAIL',
        'WEIGHT', 'PRICE', 'DPOST', 'DPOST PRICE', 'COD DETAIL MIN', 'COD DETAIL MAX',
        'COD DETAIL VOLUME', 'COD DETAIL COD', 'COD DETAIL VALUE', 'COD DETAIL QTY',
        'AMUNT', 'TYPE OF PAYMENT', 'COMMENT'
    ]
    
    rows = []
    for idx, rec in enumerate(all_records, 1):
        row_data = {col: "" for col in columns}
        row_data['NO'] = idx
        row_data['REF NO'] = rec.get('REF NO', '')
        row_data['SHIPPER NAME'] = rec.get('SHIPPER NAME', '')
        row_data['SHIPPER ADDRESS'] = rec.get('SHIPPER ADDRESS', '')
        row_data['SHIPPER AMPHUR'] = rec.get('SHIPPER AMPHUR', '')
        row_data['SHIPPER PROVINCE'] = rec.get('SHIPPER PROVINCE', '')
        row_data['SHIPPER ZIPCODE'] = rec.get('SHIPPER ZIPCODE', '')
        row_data['SHIPPER TEL'] = rec.get('SHIPPER TEL', '')
        
        row_data['RECEIVER'] = rec.get('RECEIVER', '')
        row_data['RECEIVER ADDRESS'] = rec.get('RECEIVER ADDRESS', '')
        row_data['RECEIVER AMPHUR'] = rec.get('RECEIVER AMPHUR', '')
        row_data['RECEIVER PROVINCE'] = rec.get('RECEIVER PROVINCE', '')
        row_data['RECEIVER ZIPCODE'] = rec.get('RECEIVER ZIPCODE', '')
        rows.append(row_data)
        
    return pd.DataFrame(rows, columns=columns)

def main():
    print(f"โปรแกรมแปลงข้อมูล DPost (Version {__version__})")
    # Find all PDFs in the current directory
    pdf_files = glob.glob("*.pdf")
    if not pdf_files:
        print("ไม่พบไฟล์ PDF ในโฟลเดอร์นี้ กรุณาใส่ไฟล์ PDF ที่ต้องการแปลงข้อมูล")
        return
        
    all_records = []
    for pdf_file in pdf_files:
        records = process_pdf(pdf_file)
        all_records.extend(records)
        
    if not all_records:
        print("ไม่สามารถสกัดข้อมูลจากไฟล์ PDF ได้")
        return
        
    # Create DataFrame using extracted function
    df = records_to_dataframe(all_records)
        
    # Output file path with current datetime suffix (YYYYMMDDHHMM)
    from datetime import datetime
    suffix = datetime.now().strftime("%Y%m%d%H%M")
    output_filename = f"dpost_import_{suffix}.xlsx"
    
    # Save with sheet_name="New Order Data"
    df.to_excel(output_filename, sheet_name="New Order Data", index=False)
    print(f"\nบันทึกข้อมูลเรียบร้อยแล้วลงไฟล์ Excel: {output_filename}")
    print(f"รวมข้อมูลทั้งหมด {len(df)} รายการ")

if __name__ == "__main__":
    main()
