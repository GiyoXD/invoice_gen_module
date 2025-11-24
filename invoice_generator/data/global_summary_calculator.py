"""
Global Summary Calculator

Calculates global summary data from processed_tables_data.
This includes total weights, pallet counts, and other cross-table summaries.
"""

from typing import Any, Dict
from ..utils.math_utils import safe_float_convert, safe_int_convert
import logging

logger = logging.getLogger(__name__)


class GlobalSummaryCalculator:
    """
    Calculates global summary data from processed_tables_data.
    
    This class extracts all global calculation logic from BuilderConfigResolver,
    providing a clean, testable, and reusable way to compute cross-table summaries.
    
    Summaries include:
    - Total net weight across all tables
    - Total gross weight across all tables
    - Total pallet count across all tables
    - Any other cross-table aggregations needed
    
    Usage:
        calculator = GlobalSummaryCalculator(invoice_data['processed_tables_data'])
        summaries = calculator.calculate_all()
        # summaries = {
        #     'total_net_weight': 8221.9081,
        #     'total_gross_weight': 9407.0,
        #     'total_pallets': 26
        # }
    """
    
    def __init__(self, processed_tables_data: Dict[str, Any]):
        """
        Initialize the calculator with processed tables data.
        
        Args:
            processed_tables_data: Dictionary of table data, keyed by table ID.
                Expected structure:
                {
                    '1': {'net': [...], 'gross': [...], 'pallet_count': [...]},
                    '2': {'net': [...], 'gross': [...], 'pallet_count': [...]},
                    ...
                }
        """
        self.processed_tables_data = processed_tables_data or {}
        self.summaries = {}
    
    def calculate_all(self) -> Dict[str, Any]:
        """
        Calculate all global summaries and return as a dictionary.
        
        Returns:
            Dictionary containing all calculated summaries:
            {
                'total_net_weight': float,
                'total_gross_weight': float,
                'total_pallets': int
            }
        """
        logger.debug(f"Calculating global summaries from {len(self.processed_tables_data)} tables")
        
        self.summaries = {
            'total_net_weight': self._calculate_total_net_weight(),
            'total_gross_weight': self._calculate_total_gross_weight(),
            'total_pallets': self._calculate_total_pallets(),
        }
        
        logger.debug(f"Global summaries calculated: {self.summaries}")
        return self.summaries
    
    def _calculate_total_net_weight(self) -> float:
        """
        Sum all net weights from all tables.
        
        Returns:
            Total net weight as float
        """
        total_net = 0.0
        
        for table_key, table_data in self.processed_tables_data.items():
            net_values = table_data.get('net', [])
            
            for val in net_values:
                total_net += safe_float_convert(val)
        
        logger.debug(f"Total net weight: {total_net}")
        return total_net
    
    def _calculate_total_gross_weight(self) -> float:
        """
        Sum all gross weights from all tables.
        
        Returns:
            Total gross weight as float
        """
        total_gross = 0.0
        
        for table_key, table_data in self.processed_tables_data.items():
            gross_values = table_data.get('gross', [])
            
            for val in gross_values:
                total_gross += safe_float_convert(val)
        
        logger.debug(f"Total gross weight: {total_gross}")
        return total_gross
    
    def _calculate_total_pallets(self) -> int:
        """
        Sum all pallet counts from all tables.
        
        Returns:
            Total pallet count as integer
        """
        total_pallets = 0
        
        for table_key, table_data in self.processed_tables_data.items():
            pallet_values = table_data.get('pallet_count', [])
            
            for val in pallet_values:
                total_pallets += safe_int_convert(val)
        
        logger.debug(f"Total pallets: {total_pallets}")
        return total_pallets
