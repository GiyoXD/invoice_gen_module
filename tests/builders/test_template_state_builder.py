import unittest
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from invoice_generator.builders.template_state_builder import TemplateStateBuilder

class TestTemplateStateBuilder(unittest.TestCase):

    def setUp(self):
        self.workbook = Workbook()
        self.worksheet = self.workbook.active

    def test_capture_header_and_restore_values_styles(self):
        # Setup initial worksheet with header content
        self.worksheet['A1'] = "Header 1"
        self.worksheet['B1'] = "Header 2"
        self.worksheet['A1'].font = Font(bold=True)
        self.worksheet['A1'].alignment = Alignment(horizontal="center")
        self.worksheet['A2'] = "Sub Header 1"
        self.worksheet['B2'] = "Sub Header 2"
        self.worksheet.merge_cells('A1:B1')
        self.worksheet.row_dimensions[1].height = 20
        self.worksheet.column_dimensions['A'].width = 15
        builder = TemplateStateBuilder(self.worksheet)

        # Simulate data starting at row 3
        header_end_row = 2
        builder.capture_header(header_end_row)

        # Create a new worksheet for restoration
        new_workbook = Workbook()
        new_worksheet = new_workbook.active

        # Simulate some data being written, pushing the footer down
        new_worksheet['A3'] = "Data Row 1"
        new_worksheet['B3'] = "Data Row 2"
        new_worksheet['A4'] = "Data Row 3"
        new_worksheet['B4'] = "Data Row 4"

        # Restore state
        builder.restore_state(new_worksheet, data_start_row=3)

        # Assertions for header values and styles
        self.assertEqual(new_worksheet['A1'].value, "Header 1")
        self.assertTrue(new_worksheet['A1'].font.bold)
        self.assertEqual(new_worksheet['B1'].value, None) # B1 should be None due to merge
        self.assertEqual(new_worksheet['A1'].alignment.horizontal, "center") # Alignment should still be there
        self.assertEqual(new_worksheet['A2'].value, "Sub Header 1")
        self.assertEqual(new_worksheet['B2'].value, "Sub Header 2")

        # Assertions for merged cells
        self.assertIn('A1:B1', [str(r) for r in new_worksheet.merged_cells.ranges])

        # Assertions for row heights and column widths
        self.assertEqual(new_worksheet.row_dimensions[1].height, 20)
        self.assertEqual(new_worksheet.column_dimensions['A'].width, 15)

    def test_capture_footer_and_restore_values_styles(self):
        # Setup initial worksheet with footer content
        # Data ends at row 9, footer starts at row 10
        self.worksheet['A10'] = "Footer 1"
        self.worksheet['B10'] = "Footer 2"
        self.worksheet['A10'].font = Font(italic=True)
        self.worksheet['A10'].alignment = Alignment(vertical="top")
        self.worksheet.merge_cells('A10:B10')
        self.worksheet.row_dimensions[10].height = 25
        self.worksheet.column_dimensions['B'].width = 20
        builder = TemplateStateBuilder(self.worksheet)

        # Simulate data ending at row 9, so footer starts after row 9
        data_end_row = 9
        builder.capture_footer(data_end_row)

        # Create a new worksheet for restoration
        new_workbook = Workbook()
        new_worksheet = new_workbook.active

        # Simulate header and data being written, pushing the footer down
        # Header (1 row) + Data (10 rows) = 11 rows
        new_worksheet['A1'] = "Header"
        for i in range(1, 11): # 10 data rows
            new_worksheet.cell(row=i+1, column=1, value=f"Data {i}")

        # Restore state
        # The data_start_row here is used to calculate the offset for footer merged cells.
        # It should be the row where data started in the *original* template.
        # In this test, we are simulating data starting at row 1 (after a single header row).
        builder.restore_state(new_worksheet, data_start_row=2) 

        # Assertions for footer values and styles (adjusted row)
        # The footer should be restored after the 11 rows (1 header + 10 data) in new_worksheet.
        # So, it should start at row 12.
        self.assertEqual(new_worksheet['A12'].value, "Footer 1")
        self.assertTrue(new_worksheet['A12'].font.italic)
        self.assertEqual(new_worksheet['B12'].value, None) # B12 should be None due to merge
        self.assertEqual(new_worksheet['A12'].alignment.vertical, "top")

        # Assertions for merged cells in footer (restored)
        self.assertIn('A12:B12', [str(r) for r in new_worksheet.merged_cells.ranges])

        # Assertions for row heights and column widths (restored)
        self.assertEqual(new_worksheet.row_dimensions[12].height, 25)
        self.assertEqual(new_worksheet.column_dimensions['B'].width, 20)


if __name__ == '__main__':
    unittest.main()