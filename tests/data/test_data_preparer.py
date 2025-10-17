import unittest
from invoice_generator.data.data_preparer import _to_numeric, parse_mapping_rules, prepare_data_rows

class TestDataPreparer(unittest.TestCase):

    def test_to_numeric(self):
        self.assertEqual(_to_numeric("123"), 123)
        self.assertEqual(_to_numeric("123.45"), 123.45)
        self.assertEqual(_to_numeric("1,234.56"), 1234.56)
        self.assertEqual(_to_numeric("abc"), "abc")
        self.assertEqual(_to_numeric(None), None)
        self.assertEqual(_to_numeric(123), 123)
        self.assertEqual(_to_numeric(123.45), 123.45)

    def test_parse_mapping_rules_simple(self):
        mapping_rules = {
            "col_po": {"id": "col_po", "key_index": 0},
            "col_item": {"id": "col_item", "key_index": 1},
        }
        column_id_map = {"col_po": 1, "col_item": 2}
        idx_to_header_map = {1: "PO", 2: "Item"}
        
        parsed_rules = parse_mapping_rules(mapping_rules, column_id_map, idx_to_header_map)
        
        self.assertIn("dynamic_mapping_rules", parsed_rules)
        self.assertEqual(len(parsed_rules["dynamic_mapping_rules"]), 2)
        self.assertEqual(parsed_rules["dynamic_mapping_rules"]["col_po"]["id"], "col_po")

    def test_prepare_data_rows_aggregation(self):
        data_source = {
            ("PO1", "Item1"): {"qty": 10},
            ("PO2", "Item2"): {"qty": 20},
        }
        dynamic_mapping_rules = {
            "col_po": {"id": "col_po", "key_index": 0},
            "col_item": {"id": "col_item", "key_index": 1},
            "col_qty": {"id": "col_qty", "value_key": "qty"},
        }
        column_id_map = {"col_po": 1, "col_item": 2, "col_qty": 3}
        idx_to_header_map = {1: "PO", 2: "Item", 3: "Qty"}
        
        prepared_data, _, _, _ = prepare_data_rows(
            data_source_type="aggregation",
            data_source=data_source,
            dynamic_mapping_rules=dynamic_mapping_rules,
            column_id_map=column_id_map,
            idx_to_header_map=idx_to_header_map,
            desc_col_idx=-1,
            num_static_labels=0,
            static_value_map={},
            DAF_mode=False,
        )
        
        self.assertEqual(len(prepared_data), 2)
        self.assertEqual(prepared_data[0][1], "PO1")
        self.assertEqual(prepared_data[0][2], "Item1")
        self.assertEqual(prepared_data[0][3], 10)

if __name__ == '__main__':
    unittest.main()