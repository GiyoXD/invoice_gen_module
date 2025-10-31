
from typing import Any, Dict, List, Optional, Tuple
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Alignment, Font, Side, Border
from openpyxl.utils import get_column_letter

from ..utils.layout import unmerge_row


from ..styling.models import StylingConfigModel

from ..styling.style_applier import apply_cell_style
from .bundle_accessor import BundleAccessor

class FooterBuilderStyler(BundleAccessor):
    """
    Builds and styles footer sections using pure bundle architecture.
    
    This class handles BOTH structural building (rows, cells, formulas, merges)
    AND styling (fonts, borders, colors, alignment) in a single efficient pass.
    
    Styling logic is delegated to the style_applier module for separation of concerns.
    Uses config bundles for input and @property decorators for frequently accessed values.
    """
    
    def __init__(
        self,
        worksheet: Worksheet,
        footer_row_num: int,
        style_config: Dict[str, Any],
        context_config: Dict[str, Any],
        data_config: Dict[str, Any]
    ):
        """
        Initialize FooterBuilder with bundle configs.
        
        Args:
            worksheet: The worksheet to build in
            footer_row_num: The row number where footer should be placed
            style_config: Bundle containing styling_config
            context_config: Bundle containing header_info, pallet_count, sheet_name, is_last_table, dynamic_desc_used
            data_config: Bundle containing sum_ranges, footer_config, all_tables_data, table_keys, mapping_rules, DAF_mode, override_total_text
        """
        # Initialize base class with common bundles
        super().__init__(
            worksheet=worksheet,
            style_config=style_config,
            context_config=context_config,
            data_config=data_config  # Pass data_config to base via kwargs
        )
        
        # Store FooterBuilder-specific attributes
        self.footer_row_num = footer_row_num
    
    # ========== Properties for Frequently Accessed Config Values ==========
    # Note: sheet_name, sheet_styling_config inherited from BundleAccessor
    
    @property
    def header_info(self) -> Dict[str, Any]:
        """Header information from context config."""
        return self.context_config.get('header_info', {})
    
    @property
    def sum_ranges(self) -> List[Tuple[int, int]]:
        """Sum ranges from data config."""
        return self.data_config.get('sum_ranges', [])
    
    @property
    def footer_config(self) -> Dict[str, Any]:
        """Footer configuration from data config."""
        return self.data_config.get('footer_config', {})
    
    @property
    def pallet_count(self) -> int:
        """Pallet count from context config."""
        return self.context_config.get('pallet_count', 0)
    
    @property
    def override_total_text(self) -> Optional[str]:
        """Override total text from data config."""
        return self.data_config.get('override_total_text')
    
    @property
    def DAF_mode(self) -> bool:
        """DAF mode flag from data config."""
        return self.data_config.get('DAF_mode', False)
    
    @property
    def all_tables_data(self) -> Optional[Dict[str, Any]]:
        """All tables data from data config."""
        return self.data_config.get('all_tables_data')
    
    @property
    def table_keys(self) -> Optional[List[str]]:
        """Table keys from data config."""
        return self.data_config.get('table_keys')
    
    @property
    def mapping_rules(self) -> Optional[Dict[str, Any]]:
        """Mapping rules from data config."""
        return self.data_config.get('mapping_rules')
    
    @property
    def is_last_table(self) -> bool:
        """Is last table flag from context config."""
        return self.context_config.get('is_last_table', False)
    
    @property
    def dynamic_desc_used(self) -> bool:
        """Dynamic description used flag from context config."""
        return self.context_config.get('dynamic_desc_used', False)

    def _apply_footer_cell_style(self, cell, col_id):
        """Apply footer cell style to a single cell."""
        context = {
            "col_id": col_id,
            "col_idx": cell.column,
            "is_footer": True
        }
        apply_cell_style(cell, self.sheet_styling_config, context)
    
    def _resolve_column_index(self, col_id, column_map_by_id: Dict[str, int]) -> Optional[int]:
        """
        Resolve a column ID to its actual column index.
        
        Handles both integer and string column IDs, with fallback to column_map_by_id lookup.
        
        Args:
            col_id: The column identifier (can be int, string representing int, or ID string)
            column_map_by_id: Map of column IDs to column indices
            
        Returns:
            The resolved column index (1-based), or None if not found
        """
        if col_id is None:
            return None
        
        # Handle integer column IDs
        if isinstance(col_id, int):
            return col_id + 1
        
        # Handle string column IDs
        if isinstance(col_id, str):
            try:
                # Try to parse as integer
                raw_index = int(col_id)
                return raw_index + 1
            except ValueError:
                # Look up in column map
                return column_map_by_id.get(col_id)
        
        return None

    def build(self) -> int:
        if not self.footer_config or self.footer_row_num <= 0:
            return -1

        try:
            current_footer_row = self.footer_row_num
            
            footer_type = self.footer_config.get("type", "regular")

            if footer_type == "regular":
                self._build_regular_footer(current_footer_row)
            elif footer_type == "grand_total":
                self._build_grand_total_footer(current_footer_row)

            # Apply row height to the footer row
            self._apply_footer_row_height(current_footer_row)
            
            current_footer_row += 1

            # Handle add-ons
            add_ons = self.footer_config.get("add_ons", [])
            if "summary" in add_ons:
                current_footer_row = self._build_summary_add_on(current_footer_row)



            return current_footer_row

        except Exception as e:
            print(f"ERROR: An error occurred during footer generation on row {self.footer_row_num}: {e}")
            return -1

    def _build_regular_footer(self, current_footer_row: int):
        """Build regular footer with TOTAL: text."""
        default_total_text = self.footer_config.get("total_text", "TOTAL:")
        self._build_footer_common(current_footer_row, default_total_text, unmerge_first=True)

    def _build_grand_total_footer(self, current_footer_row: int):
        """Build grand total footer with TOTAL OF: text."""
        self._build_footer_common(current_footer_row, "TOTAL OF:", unmerge_first=False)
    
    def _build_footer_common(self, current_footer_row: int, default_total_text: str, unmerge_first: bool = True):
        """
        Common footer building logic for both regular and grand total footers.
        
        Args:
            current_footer_row: The row to build the footer in
            default_total_text: Default text to use for total label
            unmerge_first: Whether to unmerge the row first
        """
        num_columns = self.header_info.get('num_columns', 1)
        column_map_by_id = self.header_info.get('column_id_map', {})

        if unmerge_first:
            unmerge_row(self.worksheet, current_footer_row, num_columns)

        # Write total text
        total_text = self.override_total_text if self.override_total_text is not None else default_total_text
        total_text_col_id = self.footer_config.get("total_text_column_id")
        total_text_col_idx = self._resolve_column_index(total_text_col_id, column_map_by_id)
        
        if total_text_col_idx:
            cell = self.worksheet.cell(row=current_footer_row, column=total_text_col_idx, value=total_text)
            self._apply_footer_cell_style(cell, total_text_col_id)

        # Write pallet count
        pallet_col_id = self.footer_config.get("pallet_count_column_id")
        pallet_col_idx = self._resolve_column_index(pallet_col_id, column_map_by_id)
        
        if pallet_col_idx and self.pallet_count > 0:
            pallet_text = f"{self.pallet_count} PALLET{'S' if self.pallet_count != 1 else ''}"
            cell = self.worksheet.cell(row=current_footer_row, column=pallet_col_idx, value=pallet_text)
            self._apply_footer_cell_style(cell, pallet_col_id)

        # Write sum formulas
        sum_column_ids = self.footer_config.get("sum_column_ids", [])
        if self.sum_ranges:
            for col_id in sum_column_ids:
                col_idx = column_map_by_id.get(col_id)
                if col_idx:
                    col_letter = get_column_letter(col_idx)
                    sum_parts = [f"{col_letter}{start}:{col_letter}{end}" for start, end in self.sum_ranges]
                    formula = f"=SUM({','.join(sum_parts)})"
                    cell = self.worksheet.cell(row=current_footer_row, column=col_idx, value=formula)
                    self._apply_footer_cell_style(cell, col_id)
        
        # Apply styling to all footer cells
        idx_to_id_map = {v: k for k, v in column_map_by_id.items()}
        for c_idx in range(1, num_columns + 1):
            cell = self.worksheet.cell(row=current_footer_row, column=c_idx)
            col_id = idx_to_id_map.get(c_idx)
            self._apply_footer_cell_style(cell, col_id)

        # Apply merge rules
        merge_rules = self.footer_config.get("merge_rules", [])
        for rule in merge_rules:
            start_column_id = rule.get("start_column_id")
            colspan = rule.get("colspan")
            resolved_start_col = self._resolve_column_index(start_column_id, column_map_by_id)
            
            if resolved_start_col and colspan:
                end_col = min(resolved_start_col + colspan - 1, num_columns)
                self.worksheet.merge_cells(start_row=current_footer_row, start_column=resolved_start_col, end_row=current_footer_row, end_column=end_col)

    def _build_summary_add_on(self, current_footer_row: int) -> int:
        from ..utils.layout import write_summary_rows # NEW IMPORT
        if self.DAF_mode and self.dynamic_desc_used and self.sheet_name == "Packing list" and self.is_last_table and self.all_tables_data and self.table_keys and self.mapping_rules:
            return write_summary_rows(
                worksheet=self.worksheet,
                start_row=current_footer_row,
                header_info=self.header_info,
                all_tables_data=self.all_tables_data,
                table_keys=self.table_keys,
                footer_config=self.footer_config,
                mapping_rules=self.mapping_rules,
                styling_config=self.sheet_styling_config,
                DAF_mode=self.DAF_mode,
                grand_total_pallets=self.pallet_count
            )
        return current_footer_row
