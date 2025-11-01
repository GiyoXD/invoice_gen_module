"""
Tests for BuilderConfigResolver class.

Tests the resolver's ability to extract and prepare configuration bundles
for specific builders from the bundled config structure.
"""
import unittest
from unittest.mock import Mock, MagicMock, patch
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from invoice_generator.config.builder_config_resolver import BuilderConfigResolver


class TestBuilderConfigResolver(unittest.TestCase):
    """Test suite for BuilderConfigResolver class."""
    
    def setUp(self):
        """Set up test fixtures before each test."""
        self.workbook = Workbook()
        self.worksheet: Worksheet = self.workbook.active
        
        # Mock config loader with sample bundled config
        self.config_loader = Mock()
        
        # Sample raw config structure (bundled v2.1)
        self.raw_config = {
            '_meta': {
                'config_version': '2.1',
                'customer': 'TEST'
            },
            'processing': {
                'sheets': ['Invoice', 'Packing'],
                'data_sources': {
                    'Invoice': 'aggregation',
                    'Packing': 'processed_tables_multi'
                }
            },
            'styling_bundle': {
                'Invoice': {
                    'rowHeights': {
                        'header': 25.0,
                        'data': 18.0,
                        'footer': 22.0
                    }
                }
            },
            'layout_bundle': {
                'Invoice': {
                    'structure': {
                        'header_row': 21,
                        'columns': [
                            {'id': 'col1', 'header': 'Item', 'width': 30},
                            {'id': 'col2', 'header': 'Quantity', 'width': 15, 'format': '#,##0'}
                        ]
                    },
                    'blanks': {
                        'add_blank_after_header': True,
                        'add_blank_before_footer': False
                    },
                    'static_content': {
                        'after_header': {'A2': 'Static content'}
                    },
                    'merge_rules': {
                        'footer': {'A1:B1': {}}
                    },
                    'data_flow': {
                        'mappings': {
                            'item_name': 'col1',
                            'quantity': 'col2'
                        }
                    },
                    'footer': {
                        'total_label': 'Total:',
                        'sum_columns': ['col2']
                    }
                }
            },
            'data_bundle': {
                'Invoice': {
                    'data_source': 'aggregation'
                }
            },
            'context': {
                'replacements': {
                    'placeholder': {
                        'INVOICE_NO': '12345'
                    }
                },
                'features': {
                    'enable_text_replacement': True
                }
            }
        }
        
        # Sample invoice data
        self.invoice_data = {
            'standard_aggregation_results': {
                'table_1': [
                    {'item_name': 'Product A', 'quantity': 10},
                    {'item_name': 'Product B', 'quantity': 20}
                ]
            },
            'custom_aggregation_results': {
                'custom_1': [{'custom_field': 'value'}]
            },
            'processed_tables_data': {
                '1': [{'field': 'data1'}],
                '2': [{'field': 'data2'}]
            }
        }
        
        # Configure mock config loader
        self.config_loader.get_raw_config.return_value = self.raw_config
        self.config_loader.get_sheet_config.return_value = {
            'data_source': 'aggregation',
            'styling_config': self.raw_config['styling_bundle']['Invoice'],
            'layout_config': self.raw_config['layout_bundle']['Invoice'],
            'data_config': self.raw_config['data_bundle']['Invoice'],
            'context_config': self.raw_config['context']
        }
        
        # Mock CLI args
        self.args = Mock(DAF=False, custom=False)
        
        # Create resolver instance
        self.resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            args=self.args,
            invoice_data=self.invoice_data,
            pallets=31
        )
    
    # ========== Initialization Tests ==========
    
    def test_init_stores_parameters(self):
        """Test that initialization stores all parameters correctly."""
        self.assertEqual(self.resolver.config_loader, self.config_loader)
        self.assertEqual(self.resolver.sheet_name, 'Invoice')
        self.assertEqual(self.resolver.worksheet, self.worksheet)
        self.assertEqual(self.resolver.args, self.args)
        self.assertEqual(self.resolver.invoice_data, self.invoice_data)
        self.assertEqual(self.resolver.pallets, 31)
    
    def test_init_caches_sheet_config(self):
        """Test that initialization caches the sheet config."""
        self.assertIsNotNone(self.resolver._sheet_config)
        self.config_loader.get_sheet_config.assert_called_once_with('Invoice')
    
    def test_init_with_context_overrides(self):
        """Test that initialization stores context overrides."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            custom_context='value'
        )
        
        self.assertEqual(resolver.context_overrides['custom_context'], 'value')
    
    # ========== Style Bundle Tests ==========
    
    def test_get_style_bundle_returns_styling_config(self):
        """Test get_style_bundle returns correct styling configuration."""
        style_bundle = self.resolver.get_style_bundle()
        
        self.assertIn('styling_config', style_bundle)
        self.assertEqual(style_bundle['styling_config'], self.raw_config['styling_bundle']['Invoice'])
    
    def test_get_style_bundle_empty_config(self):
        """Test get_style_bundle handles missing styling config."""
        self.resolver._sheet_config = {}
        style_bundle = self.resolver.get_style_bundle()
        
        self.assertEqual(style_bundle['styling_config'], {})
    
    # ========== Context Bundle Tests ==========
    
    def test_get_context_bundle_returns_base_context(self):
        """Test get_context_bundle returns correct base context."""
        context_bundle = self.resolver.get_context_bundle()
        
        self.assertEqual(context_bundle['sheet_name'], 'Invoice')
        self.assertEqual(context_bundle['args'], self.args)
        self.assertEqual(context_bundle['pallets'], 31)
        self.assertIn('all_sheet_configs', context_bundle)
    
    def test_get_context_bundle_with_additional_context(self):
        """Test get_context_bundle merges additional context."""
        context_bundle = self.resolver.get_context_bundle(
            custom_key='custom_value',
            pallet_count=50
        )
        
        self.assertEqual(context_bundle['custom_key'], 'custom_value')
        self.assertEqual(context_bundle['pallet_count'], 50)
        # Original values should still be present
        self.assertEqual(context_bundle['sheet_name'], 'Invoice')
    
    def test_get_context_bundle_applies_overrides(self):
        """Test get_context_bundle applies context overrides from init."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            override_pallets=100
        )
        
        context_bundle = resolver.get_context_bundle()
        self.assertEqual(context_bundle['override_pallets'], 100)
    
    # ========== Layout Bundle Tests ==========
    
    def test_get_layout_bundle_returns_layout_config(self):
        """Test get_layout_bundle returns complete layout configuration."""
        layout_bundle = self.resolver.get_layout_bundle()
        
        self.assertIn('sheet_config', layout_bundle)
        self.assertIn('blanks', layout_bundle)
        self.assertIn('static_content', layout_bundle)
        self.assertIn('merge_rules', layout_bundle)
    
    def test_get_layout_bundle_extracts_sub_configs(self):
        """Test get_layout_bundle extracts sub-configs correctly."""
        layout_bundle = self.resolver.get_layout_bundle()
        
        # Check blanks
        self.assertTrue(layout_bundle['blanks']['add_blank_after_header'])
        
        # Check static_content
        self.assertIn('after_header', layout_bundle['static_content'])
        
        # Check merge_rules
        self.assertIn('footer', layout_bundle['merge_rules'])
    
    def test_get_layout_bundle_empty_config(self):
        """Test get_layout_bundle handles missing layout config."""
        self.resolver._sheet_config = {}
        layout_bundle = self.resolver.get_layout_bundle()
        
        self.assertEqual(layout_bundle['blanks'], {})
        self.assertEqual(layout_bundle['static_content'], {})
        self.assertEqual(layout_bundle['merge_rules'], {})
    
    # ========== Data Bundle Tests ==========
    
    def test_get_data_bundle_returns_data_source(self):
        """Test get_data_bundle returns correct data source."""
        data_bundle = self.resolver.get_data_bundle()
        
        self.assertIn('data_source', data_bundle)
        self.assertEqual(data_bundle['data_source_type'], 'aggregation')
    
    def test_get_data_bundle_constructs_header_info(self):
        """Test get_data_bundle constructs header_info correctly."""
        data_bundle = self.resolver.get_data_bundle()
        
        header_info = data_bundle['header_info']
        self.assertIn('column_map', header_info)
        self.assertIn('column_id_map', header_info)
        self.assertEqual(header_info['column_map']['Item'], 1)
        self.assertEqual(header_info['column_map']['Quantity'], 2)
        self.assertEqual(header_info['column_id_map']['col1'], 1)
        self.assertEqual(header_info['column_id_map']['col2'], 2)
    
    def test_get_data_bundle_extracts_mapping_rules(self):
        """Test get_data_bundle extracts mapping rules."""
        data_bundle = self.resolver.get_data_bundle()
        
        mapping_rules = data_bundle['mapping_rules']
        self.assertEqual(mapping_rules['item_name'], 'col1')
        self.assertEqual(mapping_rules['quantity'], 'col2')
    
    def test_get_data_bundle_with_table_key(self):
        """Test get_data_bundle extracts specific table data."""
        # Use processed_tables data source
        self.resolver._sheet_config = {
            'data_source': 'processed_tables_multi',
            'layout_config': self.raw_config['layout_bundle']['Invoice']
        }
        
        data_bundle = self.resolver.get_data_bundle(table_key='1')
        
        # Should extract table '1' data
        self.assertEqual(data_bundle['table_key'], '1')
        # Data source should be for that specific table
        self.assertEqual(data_bundle['data_source'], [{'field': 'data1'}])
    
    def test_get_data_bundle_column_formats(self):
        """Test get_data_bundle includes column formats."""
        data_bundle = self.resolver.get_data_bundle()
        
        header_info = data_bundle['header_info']
        self.assertIn('column_formats', header_info)
        self.assertEqual(header_info['column_formats']['col2'], '#,##0')
    
    # ========== Header Bundles Tests ==========
    
    def test_get_header_bundles_returns_three_bundles(self):
        """Test get_header_bundles returns style, context, and layout bundles."""
        style_config, context_config, layout_config = self.resolver.get_header_bundles()
        
        self.assertIsInstance(style_config, dict)
        self.assertIsInstance(context_config, dict)
        self.assertIsInstance(layout_config, dict)
    
    def test_get_header_bundles_contains_expected_keys(self):
        """Test get_header_bundles contains all expected keys."""
        style_config, context_config, layout_config = self.resolver.get_header_bundles()
        
        self.assertIn('styling_config', style_config)
        self.assertIn('sheet_name', context_config)
        self.assertIn('sheet_config', layout_config)
    
    # ========== DataTable Bundles Tests ==========
    
    def test_get_datatable_bundles_returns_four_bundles(self):
        """Test get_datatable_bundles returns all four bundles."""
        bundles = self.resolver.get_datatable_bundles()
        
        self.assertEqual(len(bundles), 4)
        style_config, context_config, layout_config, data_config = bundles
        
        self.assertIsInstance(style_config, dict)
        self.assertIsInstance(context_config, dict)
        self.assertIsInstance(layout_config, dict)
        self.assertIsInstance(data_config, dict)
    
    def test_get_datatable_bundles_with_table_key(self):
        """Test get_datatable_bundles passes table_key to data bundle."""
        style_config, context_config, layout_config, data_config = self.resolver.get_datatable_bundles(table_key='table_1')
        
        self.assertEqual(data_config['table_key'], 'table_1')
    
    def test_get_datatable_bundles_contains_all_data(self):
        """Test get_datatable_bundles contains complete data."""
        style_config, context_config, layout_config, data_config = self.resolver.get_datatable_bundles()
        
        # Check data bundle has all needed keys
        self.assertIn('data_source', data_config)
        self.assertIn('header_info', data_config)
        self.assertIn('mapping_rules', data_config)
    
    # ========== Footer Bundles Tests ==========
    
    def test_get_footer_bundles_returns_three_bundles(self):
        """Test get_footer_bundles returns style, context, and data bundles."""
        bundles = self.resolver.get_footer_bundles()
        
        self.assertEqual(len(bundles), 3)
        style_config, context_config, data_config = bundles
        
        self.assertIsInstance(style_config, dict)
        self.assertIsInstance(context_config, dict)
        self.assertIsInstance(data_config, dict)
    
    def test_get_footer_bundles_with_sum_ranges(self):
        """Test get_footer_bundles includes sum_ranges in data config."""
        sum_ranges = [('A', 5, 10), ('B', 5, 10)]
        style_config, context_config, data_config = self.resolver.get_footer_bundles(sum_ranges=sum_ranges)
        
        self.assertEqual(data_config['sum_ranges'], sum_ranges)
    
    def test_get_footer_bundles_with_pallet_count(self):
        """Test get_footer_bundles uses provided pallet_count."""
        style_config, context_config, data_config = self.resolver.get_footer_bundles(pallet_count=50)
        
        self.assertEqual(context_config['pallet_count'], 50)
    
    def test_get_footer_bundles_defaults_to_resolver_pallets(self):
        """Test get_footer_bundles defaults to resolver's pallet count."""
        style_config, context_config, data_config = self.resolver.get_footer_bundles()
        
        self.assertEqual(context_config['pallet_count'], 31)
    
    def test_get_footer_bundles_with_is_last_table_flag(self):
        """Test get_footer_bundles includes is_last_table flag."""
        style_config, context_config, data_config = self.resolver.get_footer_bundles(is_last_table=True)
        
        self.assertTrue(context_config['is_last_table'])
    
    def test_get_footer_bundles_includes_footer_config(self):
        """Test get_footer_bundles includes footer config from layout."""
        style_config, context_config, data_config = self.resolver.get_footer_bundles()
        
        self.assertIn('footer_config', data_config)
        self.assertEqual(data_config['footer_config']['total_label'], 'Total:')
    
    def test_get_footer_bundles_includes_daf_mode(self):
        """Test get_footer_bundles includes DAF mode flag."""
        style_config, context_config, data_config = self.resolver.get_footer_bundles()
        
        self.assertIn('DAF_mode', data_config)
        self.assertFalse(data_config['DAF_mode'])
    
    def test_get_footer_bundles_daf_mode_true(self):
        """Test get_footer_bundles DAF mode when args.DAF is True."""
        self.args.DAF = True
        style_config, context_config, data_config = self.resolver.get_footer_bundles()
        
        self.assertTrue(data_config['DAF_mode'])
    
    # ========== Helper Method Tests ==========
    
    def test_construct_header_info_builds_column_map(self):
        """Test _construct_header_info builds correct column map."""
        layout_config = self.raw_config['layout_bundle']['Invoice']
        header_info = self.resolver._construct_header_info(layout_config)
        
        self.assertEqual(header_info['column_map']['Item'], 1)
        self.assertEqual(header_info['column_map']['Quantity'], 2)
    
    def test_construct_header_info_builds_column_id_map(self):
        """Test _construct_header_info builds correct column ID map."""
        layout_config = self.raw_config['layout_bundle']['Invoice']
        header_info = self.resolver._construct_header_info(layout_config)
        
        self.assertEqual(header_info['column_id_map']['col1'], 1)
        self.assertEqual(header_info['column_id_map']['col2'], 2)
    
    def test_construct_header_info_calculates_second_row_index(self):
        """Test _construct_header_info calculates correct second_row_index."""
        layout_config = self.raw_config['layout_bundle']['Invoice']
        header_info = self.resolver._construct_header_info(layout_config)
        
        # header_row is 21, so second_row_index should be 22
        self.assertEqual(header_info['second_row_index'], 22)
    
    def test_construct_header_info_includes_num_columns(self):
        """Test _construct_header_info includes correct number of columns."""
        layout_config = self.raw_config['layout_bundle']['Invoice']
        header_info = self.resolver._construct_header_info(layout_config)
        
        self.assertEqual(header_info['num_columns'], 2)
    
    def test_construct_header_info_empty_columns(self):
        """Test _construct_header_info handles empty columns list."""
        layout_config = {'structure': {'columns': [], 'header_row': 1}}
        header_info = self.resolver._construct_header_info(layout_config)
        
        self.assertEqual(header_info['column_map'], {})
        self.assertEqual(header_info['num_columns'], 0)
    
    def test_get_data_source_for_type_aggregation(self):
        """Test _get_data_source_for_type returns correct data for aggregation."""
        data = self.resolver._get_data_source_for_type('aggregation')
        
        self.assertEqual(data, self.invoice_data['standard_aggregation_results'])
    
    def test_get_data_source_for_type_custom_aggregation(self):
        """Test _get_data_source_for_type returns correct data for custom_aggregation."""
        data = self.resolver._get_data_source_for_type('custom_aggregation')
        
        self.assertEqual(data, self.invoice_data['custom_aggregation_results'])
    
    def test_get_data_source_for_type_processed_tables(self):
        """Test _get_data_source_for_type returns correct data for processed_tables."""
        data = self.resolver._get_data_source_for_type('processed_tables_multi')
        
        self.assertEqual(data, self.invoice_data['processed_tables_data'])
    
    def test_get_data_source_for_type_no_invoice_data(self):
        """Test _get_data_source_for_type handles missing invoice_data."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Invoice',
            worksheet=self.worksheet,
            invoice_data=None
        )
        
        data = resolver._get_data_source_for_type('aggregation')
        self.assertEqual(data, {})
    
    def test_get_data_source_for_type_unknown_type(self):
        """Test _get_data_source_for_type defaults to standard_aggregation for unknown types."""
        data = self.resolver._get_data_source_for_type('unknown_type')
        
        self.assertEqual(data, self.invoice_data['standard_aggregation_results'])
    
    def test_get_all_sheet_configs(self):
        """Test get_all_sheet_configs returns all sheet configurations."""
        all_configs = self.resolver.get_all_sheet_configs()
        
        self.assertIsInstance(all_configs, dict)
        # Should return data_bundle from raw config
        self.assertEqual(all_configs, self.raw_config['data_bundle'])


class TestBuilderConfigResolverEdgeCases(unittest.TestCase):
    """Test edge cases and error handling for BuilderConfigResolver."""
    
    def setUp(self):
        """Set up minimal test fixtures."""
        self.workbook = Workbook()
        self.worksheet = self.workbook.active
        
        # Minimal config loader
        self.config_loader = Mock()
        self.config_loader.get_raw_config.return_value = {
            'data_bundle': {},
            'styling_bundle': {},
            'layout_bundle': {}
        }
        self.config_loader.get_sheet_config.return_value = {}
    
    def test_resolver_with_no_args(self):
        """Test resolver works without args."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Test',
            worksheet=self.worksheet,
            args=None
        )
        
        context = resolver.get_context_bundle()
        self.assertIsNone(context['args'])
    
    def test_resolver_with_no_pallets(self):
        """Test resolver works with zero pallets."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Test',
            worksheet=self.worksheet,
            pallets=0
        )
        
        context = resolver.get_context_bundle()
        self.assertEqual(context['pallets'], 0)
    
    def test_get_footer_bundles_with_none_pallet_count(self):
        """Test get_footer_bundles handles None pallet_count."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Test',
            worksheet=self.worksheet,
            pallets=25
        )
        
        style_config, context_config, data_config = resolver.get_footer_bundles(pallet_count=None)
        
        # Should fall back to resolver's pallets
        self.assertEqual(context_config['pallet_count'], 25)
    
    def test_construct_header_info_with_missing_structure(self):
        """Test _construct_header_info handles missing structure."""
        resolver = BuilderConfigResolver(
            config_loader=self.config_loader,
            sheet_name='Test',
            worksheet=self.worksheet
        )
        
        header_info = resolver._construct_header_info({})
        
        self.assertEqual(header_info['column_map'], {})
        self.assertEqual(header_info['num_columns'], 0)
        self.assertEqual(header_info['second_row_index'], 2)  # default header_row=1, so 1+1=2


if __name__ == '__main__':
    unittest.main()
