"""
Test to verify that fallback configuration flows correctly through the bundler.

This test ensures that:
1. BuilderConfigResolver can access mapping rules
2. TableDataAdapter properly converts bundled format to legacy format
3. Fallback fields (fallback_on_none, fallback_on_DAF) are preserved during conversion
"""

import logging
from pathlib import Path
from invoice_generator.config.config_loader import BundledConfigLoader
from invoice_generator.config.builder_config_resolver import BuilderConfigResolver
from invoice_generator.config.table_value_adapter import TableDataAdapter

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


def test_fallback_in_bundler():
    """Test that fallback configuration flows through the bundler correctly."""
    
    # Load config using the correct config_loader
    config_path = Path("invoice_generator/config_bundled/JF_config/JF_config.json")
    config_loader = BundledConfigLoader(config_path)
    
    # Test Invoice sheet (aggregation data source)
    logger.info("\n=== Testing Invoice Sheet (aggregation) ===")
    resolver = BuilderConfigResolver(
        config_loader=config_loader,
        sheet_name="Invoice",
        worksheet=None,  # Not needed for this test
        args=None,
        invoice_data={"standard_aggregation_results": {}},
        pallets=0
    )
    
    # Get data bundle which contains mapping_rules
    data_bundle = resolver.get_data_bundle()
    mapping_rules = data_bundle.get('mapping_rules', {})
    
    logger.info(f"Found {len(mapping_rules)} mapping rules for Invoice sheet")
    
    # Check if description field has fallback
    desc_mapping = mapping_rules.get('description')
    if desc_mapping:
        logger.info(f"Description mapping: {desc_mapping}")
        has_fallback_on_none = 'fallback_on_none' in desc_mapping
        has_fallback_on_DAF = 'fallback_on_DAF' in desc_mapping
        has_fallback = 'fallback' in desc_mapping
        
        logger.info(f"  - fallback_on_none: {has_fallback_on_none} = {desc_mapping.get('fallback_on_none')}")
        logger.info(f"  - fallback_on_DAF: {has_fallback_on_DAF} = {desc_mapping.get('fallback_on_DAF')}")
        logger.info(f"  - fallback: {has_fallback} = {desc_mapping.get('fallback')}")
        
        assert has_fallback_on_none or has_fallback_on_DAF or has_fallback, \
            "Description field must have at least one fallback configuration"
    else:
        logger.warning("No description mapping found in Invoice sheet!")
    
    # Test with TableDataAdapter to verify conversion
    logger.info("\n=== Testing TableDataAdapter Conversion ===")
    header_info = {
        'column_id_map': {
            'col_po': 1,
            'col_item': 2,
            'col_desc': 3,
            'col_qty_sf': 4,
            'col_unit_price': 5,
            'col_amount': 6
        },
        'column_map': {
            'P/O NUMBER': 1,
            'ITEM NUMBER': 2,
            'DESCRIPTION': 3,
            'QUANTITY (SQFT)': 4,
            'UNIT PRICE (USD)': 5,
            'AMOUNT (USD)': 6
        }
    }
    
    adapter = TableDataAdapter(
        data_source_type='aggregation',
        data_source={},
        mapping_rules=mapping_rules,
        header_info=header_info,
        DAF_mode=False
    )
    
    # Access the converted rules
    converted_rules = adapter._convert_bundled_mappings_to_legacy(mapping_rules)
    logger.info(f"Converted {len(converted_rules)} mapping rules")
    
    # Check if fallback fields survived conversion
    if 'description' in converted_rules:
        converted_desc = converted_rules['description']
        logger.info(f"Converted description mapping: {converted_desc}")
        
        has_fallback_on_none = 'fallback_on_none' in converted_desc
        has_fallback_on_DAF = 'fallback_on_DAF' in converted_desc
        
        logger.info(f"  - fallback_on_none preserved: {has_fallback_on_none}")
        logger.info(f"  - fallback_on_DAF preserved: {has_fallback_on_DAF}")
        
        assert has_fallback_on_none, "fallback_on_none should be preserved after conversion"
        assert has_fallback_on_DAF, "fallback_on_DAF should be preserved after conversion"
    
    # Test Packing list sheet (processed_tables_multi data source)
    logger.info("\n=== Testing Packing List Sheet (processed_tables_multi) ===")
    resolver_pl = BuilderConfigResolver(
        config_loader=config_loader,
        sheet_name="Packing list",
        worksheet=None,
        args=None,
        invoice_data={"processed_tables_multi": {}},
        pallets=0
    )
    
    data_bundle_pl = resolver_pl.get_data_bundle(table_key='1')
    mapping_rules_pl = data_bundle_pl.get('mapping_rules', {})
    
    logger.info(f"Found {len(mapping_rules_pl)} mapping rules for Packing list sheet")
    
    # Check if description field has fallback
    desc_mapping_pl = mapping_rules_pl.get('description')
    if desc_mapping_pl:
        logger.info(f"Description mapping: {desc_mapping_pl}")
        has_fallback_on_none = 'fallback_on_none' in desc_mapping_pl
        has_fallback_on_DAF = 'fallback_on_DAF' in desc_mapping_pl
        
        logger.info(f"  - fallback_on_none: {has_fallback_on_none} = {desc_mapping_pl.get('fallback_on_none')}")
        logger.info(f"  - fallback_on_DAF: {has_fallback_on_DAF} = {desc_mapping_pl.get('fallback_on_DAF')}")
        
        assert has_fallback_on_none or has_fallback_on_DAF, \
            "Description field must have fallback configuration"
    
    logger.info("\n✅ All fallback configuration tests passed!")


if __name__ == "__main__":
    test_fallback_in_bundler()
