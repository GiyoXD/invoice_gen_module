import unittest
from openpyxl import Workbook
from invoice_generator.builders.template_state_builder import TemplateStateBuilder
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

class TestTemplateStateBuilder(unittest.TestCase):

    def setUp(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        self.worksheet.title = "TestSheet"

        # Create a dummy template with header, data area, and footer
        # Header (rows 1-3)
        self.worksheet['A1'] = "Header 1"
        self.worksheet['A2'] = "Header 2"
        self.worksheet['B2'].font = Font(bold=True)
        self.worksheet['A3'] = "Header 3"

        # Data area (implicitly after header, before footer)
        # For this test, we'll assume data ends at row 5

        # Footer (rows 6-7)
        self.worksheet['A6'] = "Footer 1"
        self.worksheet['C6'].fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        self.worksheet['A7'] = "Footer 2"
        self.worksheet['B7'].border = Border(left=Side(style='thin'))

        self.num_header_cols = 5 # Example value

    def test_capture_footer_rows(self):
        builder = TemplateStateBuilder(self.worksheet, self.num_header_cols)
        
        # Capture header up to row 3
        builder.capture_header(end_row=3)

        # Capture footer, assuming data ends at row 5
        builder.capture_footer(data_end_row=5)

        self.assertEqual(builder.template_footer_start_row, 6, "Footer start row should be 6")
        self.assertEqual(builder.template_footer_end_row, 7, "Footer end row should be 7")

if __name__ == '__main__':
    unittest.main()