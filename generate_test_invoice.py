import subprocess
import sys
import os

# Define the paths relative to the script's location
script_dir = os.path.dirname(__file__)
data_file = os.path.join(script_dir, "invoice_gen", "data", "CLW.json")
output_file = os.path.join(script_dir, "result_test.xlsx")
template_dir = os.path.join(script_dir, "invoice_gen", "TEMPLATE")
config_dir = os.path.join(script_dir, "invoice_gen", "config")
orchestrator_path = os.path.join(script_dir, "invoice_generator", "generate_invoice.py")

command = [
    sys.executable,
    "-m",
    "invoice_generator.generate_invoice",
    data_file,
    "--output", output_file,
    "--templatedir", template_dir,
    "--configdir", config_dir,
    "--DAF"
]

print(f"Running command: {' '.join(command)}")
try:
    subprocess.run(command, check=True)
    print(f"Invoice generated successfully at: {output_file}")
except subprocess.CalledProcessError as e:
    print(f"Error generating invoice: {e}")
except FileNotFoundError:
    print(f"Error: Python executable or script not found. Command: {' '.join(command)}")
except Exception as e:
    print(f"An unexpected error occurred: {e}")
