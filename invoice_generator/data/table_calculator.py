import logging
from typing import Any, Dict, List, Optional, Tuple, Union
from decimal import Decimal

from invoice_generator.styling.models import FooterData

logger = logging.getLogger(__name__)

class TableCalculator:
    """
    Calculates summary data (weights, pallets, leather types) from resolved table data.
    
    This class extracts business logic from the DataTableBuilder, allowing for
    separation of calculation and rendering.
    """
    
    def __init__(self, header_info: Dict[str, Any]):
        """
        Initialize the calculator.
        
        Args:
            header_info: Header information with column maps.
        """
        self.header_info = header_info
        self.col_id_map = header_info.get('column_id_map', {})
        self.idx_to_id_map = {v: k for k, v in self.col_id_map.items()}
        
        # Initialize summaries
        self.leather_summary = {
            'BUFFALO': {'pallet_count': 0},
            'COW': {'pallet_count': 0}
        }
        self.weight_summary = {
            'net': 0.0,
            'gross': 0.0
        }
        self.total_pallets = 0

    def calculate(self, resolved_data: Dict[str, Any]) -> FooterData:
        """
        Perform all calculations on the provided data.
        
        Args:
            resolved_data: The data prepared by TableDataAdapter.
            
        Returns:
            FooterData object containing all calculated summaries.
        """
        data_rows = resolved_data.get('data_rows', [])
        pallet_counts = resolved_data.get('pallet_counts', [])
        
        # Calculate total pallets
        self.total_pallets = sum(int(p) for p in pallet_counts if p is not None and str(p).isdigit())
        
        # Process each row
        for i, row_data in enumerate(data_rows):
            self._update_weight_summary(row_data)
            self._update_leather_summary(row_data, i, pallet_counts)
            
        # Determine row indices (logic moved from DataTableBuilder)
        num_columns = self.header_info.get('num_columns', 0)
        data_writing_start_row = self.header_info.get('second_row_index', 0) + 1
        actual_rows_to_process = len(data_rows)
        
        data_start_row = data_writing_start_row
        data_end_row = data_start_row + actual_rows_to_process - 1 if actual_rows_to_process > 0 else data_start_row - 1
        footer_row_final = data_end_row + 1
        
        return FooterData(
            footer_row_start_idx=footer_row_final,
            data_start_row=data_start_row,
            data_end_row=data_end_row,
            total_pallets=self.total_pallets,
            leather_summary=self.leather_summary,
            weight_summary=self.weight_summary
        )

    def _update_weight_summary(self, row_data: Dict[int, Any]):
        """Updates the running totals for Net and Gross weight."""
        net_col_idx = self.col_id_map.get('col_net_weight')
        gross_col_idx = self.col_id_map.get('col_gross_weight')
        
        if net_col_idx and net_col_idx in row_data:
            try:
                val = row_data[net_col_idx]
                if isinstance(val, (int, float)):
                    self.weight_summary['net'] += float(val)
                elif isinstance(val, str) and val.replace('.', '', 1).isdigit():
                    self.weight_summary['net'] += float(val)
            except (ValueError, TypeError):
                pass
                
        if gross_col_idx and gross_col_idx in row_data:
            try:
                val = row_data[gross_col_idx]
                if isinstance(val, (int, float)):
                    self.weight_summary['gross'] += float(val)
                elif isinstance(val, str) and val.replace('.', '', 1).isdigit():
                    self.weight_summary['gross'] += float(val)
            except (ValueError, TypeError):
                pass

    def _update_leather_summary(self, row_data: Dict[int, Any], row_index: int, pallet_counts: List[Any]):
        """Updates the running totals for Buffalo and Cow leather."""
        desc_col_idx = self.col_id_map.get('col_desc')
        if not desc_col_idx:
            return

        description = str(row_data.get(desc_col_idx, "")).upper()
        
        if "BUFFALO" in description:
            target_type = 'BUFFALO'
        else:
            target_type = 'COW'
            
        if target_type:
            # Add pallet count for this row
            if row_index < len(pallet_counts):
                pallet_val = pallet_counts[row_index]
                if pallet_val is not None and str(pallet_val).replace('.', '', 1).isdigit():
                    self.leather_summary[target_type]['pallet_count'] += int(float(pallet_val))
            
            # Sum numeric columns
            for col_idx, value in row_data.items():
                col_id = self.idx_to_id_map.get(col_idx)
                if not col_id or col_id == 'col_desc':
                    continue
                
                try:
                    if isinstance(value, (int, float)):
                        num_val = value
                    elif isinstance(value, str) and value.replace('.', '', 1).isdigit():
                         num_val = float(value)
                    else:
                        continue
                        
                    if col_id not in self.leather_summary[target_type]:
                        self.leather_summary[target_type][col_id] = 0
                        
                    self.leather_summary[target_type][col_id] += num_val
                except (ValueError, TypeError):
                    continue
