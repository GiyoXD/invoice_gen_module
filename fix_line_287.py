with open(r'c:\Users\JPZ031127\Desktop\refactor\invoice_generator\builders\layout_builder.py', 'r', encoding='utf-8') as f:
    lines = f.readlines()

# Fix line 287: Remove data_source_type check in pallet_count logic
lines[284] = "            # Prepare footer parameters\n"
lines[285] = "            # Use local_chunk_pallets from data if available, otherwise use grand total\n"
lines[286] = "            if local_chunk_pallets > 0:\n"
lines[287] = "                pallet_count = local_chunk_pallets\n"
lines[288] = "            else:\n"
lines[289] = "                pallet_count = self.final_grand_total_pallets\n"

# Fix line 330: Use self.args.DAF instead of data_source_type check
lines[329] = "                'DAF_mode': self.args.DAF if self.args and hasattr(self.args, 'DAF') else False,\n"

with open(r'c:\Users\JPZ031127\Desktop\refactor\invoice_generator\builders\layout_builder.py', 'w', encoding='utf-8') as f:
    f.writelines(lines)

print("Fixed lines 287 and 330 - removed data_source_type checks")
