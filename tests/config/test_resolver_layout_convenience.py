"""
Tests for the new get_layout_bundles_with_data() convenience method.
"""
import unittest
from unittest.mock import Mock
from openpyxl import Workbook
from pathlib import Path

from invoice_generator.config.config_loader import BundledConfigLoader
from invoice_generator.config.builder_config_resolver import BuilderConfigResolver


class TestLayoutBundlesWithData(unittest.TestCase):
    """Test the convenience method for LayoutBuilder bundle preparation."""
    
    @classmethod
    def setUpClass(cls):
        """Load the real JF config once."""
        config_path = Path(__file__).parent.parent.parent / "invoice_generator" / "config_bundled" / "JF_config" / "JF_config.json"
        cls.config_loader = BundledConfigLoader(config_path)
    
    def setUp(self):
        """Set up test fixtures."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        
        self.invoice_data = {
            'standard_aggregation_results': {
                'table_1': [
                    {'po': 'PO123', 'item': 'ITEM001', 'sqft': 100.50, 'unit_price': 5.25}
                ]
            }
        }
        
        self.args = Mock(DAF=False, custom=False)
    
    def test_get_layout_bundles_with_data_returns_three_bundles(self):
        """Test that get_layout_bundles_with_data returns 3 bundles."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data
        )
        
        bundles = resolver.get_layout_bundles_with_data()
        
        self.assertEqual(len(bundles), 3)
        style_config, context_config, layout_config = bundles
        
        self.assertIsInstance(style_config, dict)
        self.assertIsInstance(context_config, dict)
        self.assertIsInstance(layout_config, dict)
    
    def test_layout_config_contains_merged_data(self):
        """Test that layout_config contains both layout AND data fields."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data
        )
        
        style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()
        
        # Should contain layout fields
        self.assertIn('sheet_config', layout_config)
        self.assertIn('blanks', layout_config)
        
        # Should ALSO contain data fields (merged from data_config)
        self.assertIn('data_source', layout_config)
        self.assertIn('data_source_type', layout_config)
        self.assertIn('header_info', layout_config)
        self.assertIn('mapping_rules', layout_config)
    
    def test_data_source_type_is_correct(self):
        """Test that data_source_type is correctly set in merged layout_config."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data
        )
        
        style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()
        
        self.assertEqual(layout_config['data_source_type'], 'aggregation')
    
    def test_header_info_is_included(self):
        """Test that header_info is properly included in layout_config."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data
        )
        
        style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()
        
        header_info = layout_config['header_info']
        self.assertIsInstance(header_info, dict)
        self.assertIn('column_map', header_info)
        self.assertIn('column_id_map', header_info)
        self.assertEqual(header_info['second_row_index'], 22)  # header_row 21 + 1
    
    def test_mapping_rules_are_included(self):
        """Test that mapping_rules are properly included in layout_config."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data
        )
        
        style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()
        
        mapping_rules = layout_config['mapping_rules']
        self.assertIsInstance(mapping_rules, dict)
        self.assertIn('po', mapping_rules)
        self.assertIn('item', mapping_rules)
    
    def test_with_table_key_extracts_correct_data(self):
        """Test that table_key parameter works correctly."""
        # Use Packing list data
        invoice_data = {
            'processed_tables_data': {
                '1': [{'po': 'PO123'}],
                '2': [{'po': 'PO456'}]
            }
        }
        
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Packing list',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=invoice_data
        )
        
        # Get data for table '1'
        style_config, context_config, layout_config = resolver.get_layout_bundles_with_data(table_key='1')
        
        data_source = layout_config['data_source']
        self.assertEqual(len(data_source), 1)
        self.assertEqual(data_source[0]['po'], 'PO123')
    
    def test_works_with_different_sheets(self):
        """Test that method works for all sheet types."""
        sheets = ['Invoice', 'Contract', 'Packing list']
        
        for sheet_name in sheets:
            with self.subTest(sheet=sheet_name):
                resolver = BuilderConfigResolver(
                    config_loader=self.config_loader,
                    sheet_name=sheet_name,
                    worksheet=self.worksheet,
                    args=self.args,
                    invoice_data=self.invoice_data
                )
                
                style_config, context_config, layout_config = resolver.get_layout_bundles_with_data()
                
                # All sheets should have merged data
                self.assertIn('data_source', layout_config)
                self.assertIn('data_source_type', layout_config)
                self.assertIn('header_info', layout_config)


if __name__ == '__main__':
    unittest.main()
