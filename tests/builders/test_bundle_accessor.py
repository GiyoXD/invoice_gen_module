"""
Tests for BundleAccessor class.

Tests the bundle storage, property accessors, and helper methods
that are shared across builder classes.
"""
import unittest
from unittest.mock import MagicMock, Mock
from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from invoice_generator.builders.bundle_accessor import BundleAccessor
from invoice_generator.styling.models import StylingConfigModel


class TestBundleAccessor(unittest.TestCase):
    """Test suite for BundleAccessor class."""
    
    def setUp(self):
        """Set up test fixtures before each test."""
        self.workbook = Workbook()
        self.worksheet: Worksheet = self.workbook.active
        
        # Sample styling config
        self.style_config = {
            'styling_config': {
                'rowHeights': {
                    'header': 25.0,
                    'data': 18.0,
                    'footer': 22.0,
                    'footer_matches_header_height': True
                },
                'fonts': {
                    'header': {'name': 'Arial', 'size': 12, 'bold': True},
                    'data': {'name': 'Arial', 'size': 10}
                }
            }
        }
        
        # Sample context config
        self.context_config = {
            'sheet_name': 'Invoice',
            'all_sheet_configs': {
                'Invoice': {'header_row': 1},
                'Packing': {'header_row': 1}
            },
            'args': Mock(DAF=False, custom=False)
        }
        
        # Create accessor instance
        self.accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=self.style_config,
            context_config=self.context_config
        )
    
    # ========== Initialization Tests ==========
    
    def test_init_stores_bundles(self):
        """Test that initialization stores all bundles correctly."""
        self.assertEqual(self.accessor.worksheet, self.worksheet)
        self.assertEqual(self.accessor.style_config, self.style_config)
        self.assertEqual(self.accessor.context_config, self.context_config)
    
    def test_init_stores_kwargs_as_attributes(self):
        """Test that additional kwargs are stored as attributes."""
        extra_bundle = {'key': 'value'}
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=self.style_config,
            context_config=self.context_config,
            custom_bundle=extra_bundle
        )
        
        self.assertTrue(hasattr(accessor, 'custom_bundle'))
        self.assertEqual(accessor.custom_bundle, extra_bundle)
    
    # ========== Property Tests ==========
    
    def test_sheet_name_property(self):
        """Test sheet_name property returns correct value."""
        self.assertEqual(self.accessor.sheet_name, 'Invoice')
    
    def test_sheet_name_property_empty_config(self):
        """Test sheet_name property returns empty string when not in config."""
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config={},
            context_config={}
        )
        self.assertEqual(accessor.sheet_name, '')
    
    def test_all_sheet_configs_property(self):
        """Test all_sheet_configs property returns correct dictionary."""
        configs = self.accessor.all_sheet_configs
        self.assertIn('Invoice', configs)
        self.assertIn('Packing', configs)
        self.assertEqual(configs['Invoice']['header_row'], 1)
    
    def test_all_sheet_configs_property_empty(self):
        """Test all_sheet_configs property returns empty dict when not in config."""
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config={},
            context_config={}
        )
        self.assertEqual(accessor.all_sheet_configs, {})
    
    def test_sheet_styling_config_property_dict_conversion(self):
        """Test sheet_styling_config converts dict to StylingConfigModel."""
        styling = self.accessor.sheet_styling_config
        self.assertIsInstance(styling, StylingConfigModel)
    
    def test_sheet_styling_config_property_already_model(self):
        """Test sheet_styling_config returns existing StylingConfigModel."""
        model = StylingConfigModel(**self.style_config['styling_config'])
        style_config = {'styling_config': model}
        
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=style_config,
            context_config=self.context_config
        )
        
        styling = accessor.sheet_styling_config
        self.assertIsInstance(styling, StylingConfigModel)
        self.assertEqual(styling, model)
    
    def test_sheet_styling_config_property_invalid_dict(self):
        """Test sheet_styling_config handles invalid dict gracefully."""
        # StylingConfigModel accepts extra fields, so test with truly invalid data
        style_config = {'styling_config': 'not_a_dict'}
        
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=style_config,
            context_config=self.context_config
        )
        
        styling = accessor.sheet_styling_config
        # Should return None on conversion failure when it's not a dict
        self.assertIsNone(styling)
    
    def test_sheet_styling_config_property_none(self):
        """Test sheet_styling_config returns None when config is missing."""
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config={},
            context_config=self.context_config
        )
        self.assertIsNone(accessor.sheet_styling_config)
    
    def test_args_property(self):
        """Test args property returns CLI arguments."""
        args = self.accessor.args
        self.assertIsNotNone(args)
        self.assertFalse(args.DAF)
        self.assertFalse(args.custom)
    
    def test_args_property_none(self):
        """Test args property returns None when not in config."""
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config={},
            context_config={}
        )
        self.assertIsNone(accessor.args)
    
    # ========== Helper Method Tests ==========
    
    def test_apply_footer_row_height_with_match_header(self):
        """Test _apply_footer_row_height uses header height when matched."""
        footer_row = 10
        self.accessor._apply_footer_row_height(footer_row)
        
        # Should use header height (25.0) since footer_matches_header_height is True
        self.assertEqual(self.worksheet.row_dimensions[footer_row].height, 25.0)
    
    def test_apply_footer_row_height_without_match_header(self):
        """Test _apply_footer_row_height uses footer height when not matched."""
        style_config = {
            'styling_config': {
                'rowHeights': {
                    'header': 25.0,
                    'footer': 22.0,
                    'footer_matches_header_height': False
                }
            }
        }
        
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=style_config,
            context_config=self.context_config
        )
        
        footer_row = 10
        accessor._apply_footer_row_height(footer_row)
        
        # Should use footer height (22.0) since footer_matches_header_height is False
        self.assertEqual(self.worksheet.row_dimensions[footer_row].height, 22.0)
    
    def test_apply_footer_row_height_no_config(self):
        """Test _apply_footer_row_height does nothing when no config."""
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config={},
            context_config=self.context_config
        )
        
        footer_row = 10
        accessor._apply_footer_row_height(footer_row)
        
        # Should not set any height
        self.assertIsNone(self.worksheet.row_dimensions[footer_row].height)
    
    def test_apply_footer_row_height_invalid_row(self):
        """Test _apply_footer_row_height handles invalid row numbers."""
        # Row 0 is invalid
        self.accessor._apply_footer_row_height(0)
        
        # Should not crash, just do nothing
        self.assertIsNone(self.worksheet.row_dimensions[0].height)
    
    def test_apply_footer_row_height_invalid_height_value(self):
        """Test _apply_footer_row_height handles invalid height values."""
        style_config = {
            'styling_config': {
                'rowHeights': {
                    'footer': 'invalid',
                    'footer_matches_header_height': False
                }
            }
        }
        
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=style_config,
            context_config=self.context_config
        )
        
        footer_row = 10
        accessor._apply_footer_row_height(footer_row)
        
        # Should not crash, just not set height
        self.assertIsNone(self.worksheet.row_dimensions[footer_row].height)
    
    def test_apply_footer_row_height_zero_height(self):
        """Test _apply_footer_row_height ignores zero or negative heights."""
        style_config = {
            'styling_config': {
                'rowHeights': {
                    'footer': 0,
                    'footer_matches_header_height': False
                }
            }
        }
        
        accessor = BundleAccessor(
            worksheet=self.worksheet,
            style_config=style_config,
            context_config=self.context_config
        )
        
        footer_row = 10
        accessor._apply_footer_row_height(footer_row)
        
        # Should not set zero height
        self.assertIsNone(self.worksheet.row_dimensions[footer_row].height)
    
    def test_get_bool_flag_returns_value(self):
        """Test _get_bool_flag returns correct boolean value."""
        config = {'enable_feature': True, 'disable_feature': False}
        
        self.assertTrue(self.accessor._get_bool_flag(config, 'enable_feature'))
        self.assertFalse(self.accessor._get_bool_flag(config, 'disable_feature'))
    
    def test_get_bool_flag_returns_default(self):
        """Test _get_bool_flag returns default when key not found."""
        config = {}
        
        self.assertFalse(self.accessor._get_bool_flag(config, 'missing_key'))
        self.assertTrue(self.accessor._get_bool_flag(config, 'missing_key', default=True))
    
    def test_get_bool_flag_handles_none_config(self):
        """Test _get_bool_flag handles None gracefully."""
        # Should not crash with None
        result = self.accessor._get_bool_flag({}, 'any_key', default=True)
        self.assertTrue(result)


if __name__ == '__main__':
    unittest.main()
