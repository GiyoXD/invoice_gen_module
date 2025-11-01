import json
from typing import Dict, Any, List, Optional
from ..styling.models import StylingConfigModel


class BundledConfigLoader:
    """
    Loader for the new bundled config format (v2.0+).
    
    Provides clean access to config sections:
    - processing: sheets and data sources
    - layout_bundle: structure, data_flow, content, footer per sheet
    - styling_bundle: styling configuration per sheet
    - features: feature flags
    - defaults: default settings
    """
    
    def __init__(self, config_data: Dict[str, Any]):
        """Initialize with loaded config data."""
        self.config_data = config_data
        self._meta = config_data.get('_meta', {})
        self.processing = config_data.get('processing', {})
        self.layout_bundle = config_data.get('layout_bundle', {})
        self.styling_bundle = config_data.get('styling_bundle', {})
        self.features = config_data.get('features', {})
        self.defaults = config_data.get('defaults', {})
        self.data_preparation_hint = config_data.get('data_preparation_module_hint', {})
    
    @property
    def config_version(self) -> str:
        """Get config version from metadata."""
        return self._meta.get('config_version', 'unknown')
    
    @property
    def customer(self) -> str:
        """Get customer name from metadata."""
        return self._meta.get('customer', '')
    
    # ========== Processing Configuration ==========
    
    def get_sheets_to_process(self) -> List[str]:
        """Get list of sheets to process."""
        return self.processing.get('sheets', [])
    
    def get_data_source(self, sheet_name: str) -> Optional[str]:
        """Get data source for a specific sheet."""
        data_sources = self.processing.get('data_sources', {})
        return data_sources.get(sheet_name)
    
    def get_sheet_data_map(self) -> Dict[str, str]:
        """Get complete sheet to data source mapping."""
        return self.processing.get('data_sources', {})
    
    # ========== Layout Configuration ==========
    
    def get_sheet_layout(self, sheet_name: str) -> Dict[str, Any]:
        """Get complete layout configuration for a sheet."""
        return self.layout_bundle.get(sheet_name, {})
    
    def get_sheet_structure(self, sheet_name: str) -> Dict[str, Any]:
        """Get structure section (start_row, columns) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('structure', {})
    
    def get_sheet_data_flow(self, sheet_name: str) -> Dict[str, Any]:
        """Get data_flow section (mappings) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('data_flow', {})
    
    def get_sheet_content(self, sheet_name: str) -> Dict[str, Any]:
        """Get content section (static content) for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('content', {})
    
    def get_sheet_footer(self, sheet_name: str) -> Dict[str, Any]:
        """Get footer section for a sheet."""
        layout = self.get_sheet_layout(sheet_name)
        return layout.get('footer', {})
    
    # ========== Styling Configuration ==========
    
    def get_sheet_styling(self, sheet_name: str) -> Dict[str, Any]:
        """Get complete styling configuration for a sheet."""
        return self.styling_bundle.get(sheet_name, {})
    
    def get_styling_defaults(self) -> Dict[str, Any]:
        """Get default styling configuration."""
        return self.styling_bundle.get('defaults', {})
    
    # ========== Data Mapping (Backward Compatibility) ==========
    
    def build_legacy_sheet_config(self, sheet_name: str) -> Dict[str, Any]:
        """
        Build a config dict in the old format for a specific sheet.
        This helps transition code that expects the old format.
        """
        layout = self.get_sheet_layout(sheet_name)
        styling = self.get_sheet_styling(sheet_name)
        
        structure = layout.get('structure', {})
        data_flow = layout.get('data_flow', {})
        content = layout.get('content', {})
        footer_config = layout.get('footer', {})
        
        columns = structure.get('columns', [])
        
        # Build header_to_write from structure.columns
        header_to_write = self._build_header_from_columns(columns)
        
        # Build a column_id to format map from structure
        column_formats = self._extract_column_formats(columns)
        
        # Build mappings from data_flow.mappings with formats merged in
        mappings = self._build_mappings_from_data_flow(
            data_flow.get('mappings', {}), 
            column_formats
        )
        
        # Build static content
        static_content_before_footer = self._build_static_content(content)
        
        # Combine into legacy format
        legacy_config = {
            'start_row': structure.get('start_row', 1),
            'header_to_write': header_to_write,
            'mappings': mappings,
            'styling': styling,
            'footer_config': footer_config,
            'add_blank_before_footer': footer_config.get('add_blank_before', False),
        }
        
        if static_content_before_footer:
            legacy_config['static_content_before_footer'] = static_content_before_footer
        
        # Add static content for columns
        if 'static' in content and 'col_static' in content['static']:
            legacy_config['static_content'] = content['static']['col_static']
        
        return legacy_config
    
    def _build_header_from_columns(self, columns: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Convert new columns format to old header_to_write format."""
        headers = []
        col_index = 0
        
        for col in columns:
            col_id = col.get('id', '')
            header_text = col.get('header', '')
            rowspan = col.get('rowspan', 1)
            colspan = col.get('colspan', 1)
            
            # Handle parent column with children (e.g., Quantity with PCS/SF)
            if 'children' in col:
                # Add parent header
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': 1,
                    'colspan': len(col['children'])
                })
                
                # Add children headers
                for child in col['children']:
                    headers.append({
                        'row': 1,
                        'col': col_index,
                        'text': child.get('header', ''),
                        'id': child.get('id', ''),
                        'rowspan': 1,
                        'colspan': 1
                    })
                    col_index += 1
            else:
                headers.append({
                    'row': 0,
                    'col': col_index,
                    'text': header_text,
                    'id': col_id,
                    'rowspan': rowspan,
                    'colspan': colspan
                })
                col_index += 1
        
        return headers
    
    def _extract_column_formats(self, columns: List[Dict[str, Any]]) -> Dict[str, str]:
        """Extract column ID to format mapping from columns structure."""
        formats = {}
        
        for col in columns:
            col_id = col.get('id', '')
            col_format = col.get('format')
            
            if col_format:
                formats[col_id] = col_format
            
            # Handle children columns
            if 'children' in col:
                for child in col['children']:
                    child_id = child.get('id', '')
                    child_format = child.get('format')
                    if child_format:
                        formats[child_id] = child_format
        
        return formats
    
    def _build_mappings_from_data_flow(
        self, 
        data_flow_mappings: Dict[str, Any],
        column_formats: Dict[str, str]
    ) -> Dict[str, Any]:
        """Convert new data_flow.mappings format to old mappings format."""
        mappings = {}
        
        for field_name, mapping in data_flow_mappings.items():
            column_id = mapping.get('column')
            
            legacy_mapping = {
                'id': column_id
            }
            
            # Handle source_key (old key_index)
            if 'source_key' in mapping:
                legacy_mapping['key_index'] = mapping['source_key']
            
            # Handle source_value (old value_key)
            if 'source_value' in mapping:
                legacy_mapping['value_key'] = mapping['source_value']
            
            # Handle formula
            if 'formula' in mapping:
                legacy_mapping['formula'] = mapping['formula']
            
            # Handle fallback
            if 'fallback' in mapping:
                legacy_mapping['fallback_on_none'] = mapping['fallback']
                legacy_mapping['fallback_on_DAF'] = mapping['fallback']
            
            # Add format from structure columns if available
            if column_id and column_id in column_formats:
                legacy_mapping['number_format'] = column_formats[column_id]
            
            mappings[field_name] = legacy_mapping
        
        return mappings
    
    def _build_static_content(self, content: Dict[str, Any]) -> Dict[str, str]:
        """Extract static content before footer."""
        static = content.get('static', {})
        before_footer = static.get('before_footer', {})
        
        if before_footer and 'col' in before_footer:
            col_index = before_footer['col']
            text = before_footer.get('text', '')
            return {str(col_index): text}
        
        return {}
    
    # ========== Features ==========
    
    def is_feature_enabled(self, feature_name: str) -> bool:
        """Check if a feature is enabled."""
        return self.features.get(feature_name, False)


def load_config(config_path: str) -> Dict[str, Any]:
    """Loads the main configuration from a JSON file."""
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


def load_bundled_config(config_path: str) -> BundledConfigLoader:
    """Loads and wraps a bundled config file."""
    config_data = load_config(config_path)
    return BundledConfigLoader(config_data)


def load_styling_config(sheet_config: Dict[str, Any]) -> StylingConfigModel:
    """Parses the styling portion of the config into Pydantic models."""
    print(f"DEBUG: sheet_config in load_styling_config: {sheet_config.get('styling', {})}")
    return StylingConfigModel(**sheet_config.get('styling', {}))
