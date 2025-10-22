import unittest
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet

from invoice_generator.builders.template_state_builder import TemplateStateBuilder

class TestTemplateStateBuilder(unittest.TestCase):

    def setUp(self):
        self.workbook = openpyxl.Workbook()
        self.worksheet: Worksheet = self.workbook.active

        # Create a dummy template with header, data, footer, and grand total footer
        # Header
        self.worksheet['A1'] = 'Header 1'
        self.worksheet['B1'] = 'Header 2'

        # Data
        self.worksheet['A2'] = 'Data 1'
        self.worksheet['B2'] = 'Data 2'

        # Footer
        self.worksheet['A4'] = 'Footer 1'
        self.worksheet['B4'] = 'Footer 2'
        self.worksheet.merge_cells('A4:B4')

        # Grand Total Footer (after a blank row)
        self.worksheet['A6'] = 'Grand Total'
        self.worksheet['B6'] = '1000'

    def test_capture_and_restore_footer_with_grand_total(self):
        # 1. Capture the template state
        template_builder = TemplateStateBuilder(self.worksheet, num_header_cols=2)
        template_builder.capture_header(end_row=1)
        template_builder.capture_footer(data_end_row=2, max_possible_footer_row=6)

        # 2. Assert that the footer and grand total footer are captured correctly
        self.assertEqual(template_builder.template_footer_start_row, 4)
        self.assertEqual(template_builder.template_footer_end_row, 6)
        self.assertEqual(len(template_builder.footer_state), 3) # Footer row, blank row, grand total row
        self.assertEqual(template_builder.footer_state[0][0]['value'], 'Footer 1')
        self.assertEqual(template_builder.footer_state[2][0]['value'], 'Grand Total')

        # 3. Create a new workbook and restore the state
        new_workbook = openpyxl.Workbook()
        new_worksheet = new_workbook.active
        template_builder.restore_state(new_worksheet, data_start_row=2, data_table_end_row=3)

        # 4. Assert that the footer and grand total footer are restored at the correct positions
        self.assertEqual(new_worksheet['A4'].value, 'Footer 1')
        self.assertIn('A4:B4', [str(r) for r in new_worksheet.merged_cells.ranges])
        self.assertEqual(new_worksheet['A6'].value, 'Grand Total')

if __name__ == '__main__':
    unittest.main()