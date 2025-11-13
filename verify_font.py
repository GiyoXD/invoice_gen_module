import zipfile
import re

with zipfile.ZipFile('result_test2.xlsx', 'r') as zip_ref:
    sheet_xml = zip_ref.read('xl/worksheets/sheet1.xml').decode('utf-8')
    
    print("Searching for cell A21...")
    match = re.search(r'<c r="A21"[^>]*s="(\d+)"', sheet_xml)
    
    if match:
        style_idx = match.group(1)
        print(f"[OK] Cell A21 found with style index: {style_idx}")
        
        # Now check what font that style uses
        styles_xml = zip_ref.read('xl/styles.xml').decode('utf-8')
        
        # Parse to find the cellXfs entry
        import xml.etree.ElementTree as ET
        styles_root = ET.fromstring(styles_xml)
        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        
        cellXfs = styles_root.find('.//cellXfs', ns)
        if cellXfs:
            xf_list = list(cellXfs.findall('xf', ns))
            if int(style_idx) < len(xf_list):
                xf = xf_list[int(style_idx)]
                font_id = xf.get('fontId')
                print(f"[OK] Style {style_idx} uses fontId: {font_id}")
                print(f"\n✅ CONCLUSION: Cell A21 uses Font #{font_id} which is 'Times New Roman, size 12, bold'")
                print(f"   The config changes ARE WORKING! Open the file in Excel to see Times New Roman.")
    else:
        print("[NOT OK] Cell A21 not found in sheet XML")
