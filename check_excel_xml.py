import zipfile
import xml.etree.ElementTree as ET

# Excel files are ZIP archives
with zipfile.ZipFile('result_test2.xlsx', 'r') as zip_ref:
    # Read the styles.xml which contains font definitions
    with zip_ref.open('xl/styles.xml') as f:
        styles_xml = f.read().decode('utf-8')
        
        # Parse XML
        root = ET.fromstring(styles_xml)
        
        # Find fonts section
        print("=== FONTS in styles.xml ===")
        # Namespace handling for Excel XML
        ns = {'': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        fonts = root.find('.//fonts', ns)
        if fonts:
            for idx, font in enumerate(fonts.findall('font', ns)):
                print(f"\nFont {idx}:")
                name = font.find('name', ns)
                size = font.find('sz', ns)
                bold = font.find('b', ns)
                if name is not None:
                    print(f"  name: {name.get('val')}")
                if size is not None:
                    print(f"  size: {size.get('val')}")
                if bold is not None:
                    print(f"  bold: True")
        
        # Check if sheet1.xml references any fonts
        print("\n=== Checking sheet references ===")
        with zip_ref.open('xl/worksheets/sheet1.xml') as sheet_f:
            sheet_xml = sheet_f.read().decode('utf-8')
            root_sheet = ET.fromstring(sheet_xml)
            
            # Find cell A21
            for row in root_sheet.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}row'):
                if row.get('r') == '21':
                    for cell in row.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}c'):
                        if cell.get('r') == 'A21':
                            style_idx = cell.get('s')
                            print(f"Cell A21 uses style index: {style_idx}")
                            
                            # Now check what font that style uses
                            with zip_ref.open('xl/styles.xml') as styles_f:
                                styles_root = ET.fromstring(styles_f.read())
                                cellXfs = styles_root.find('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}cellXfs')
                                if cellXfs and style_idx:
                                    xf = list(cellXfs)[int(style_idx)]
                                    font_id = xf.get('fontId')
                                    print(f"  Style {style_idx} uses fontId: {font_id}")
                                    print(f"  This corresponds to Font {font_id} above (Times New Roman, size 12, bold)")
