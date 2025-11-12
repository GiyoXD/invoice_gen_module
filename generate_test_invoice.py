import subprocess
import sys
import os
import argparse

# Define the paths relative to the script's location
script_dir = os.path.dirname(__file__)
data_file = os.path.join(script_dir, "invoice_generator", "JF.json")  # Changed to JF (has bundled config)
output_file = os.path.join(script_dir, "result_test2.xlsx")
template_dir = os.path.join(script_dir, "invoice_generator", "template")
config_dir = os.path.join(script_dir, "invoice_generator", "config_bundled")  # Changed to bundled config
orchestrator_path = os.path.join(script_dir, "invoice_generator", "generate_invoice.py")

# Parse command-line arguments for logging control
parser = argparse.ArgumentParser(description="Test wrapper for invoice generation")
parser.add_argument("--log-level", choices=['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'],
                    default='INFO', help="Set logging level (default: INFO)")
parser.add_argument("--debug", action="store_true", help="Enable debug logging")
parser.add_argument("--no-DAF", dest="DAF", action="store_false", help="Disable DAF mode")
parser.add_argument("--DAF", dest="DAF", action="store_true", help="Enable DAF mode (default)")
parser.set_defaults(DAF=True)
args = parser.parse_args()

command = [
    sys.executable,
    "-m",
    "invoice_generator.generate_invoice",
    data_file,
    "--output", output_file,
    "--templatedir", template_dir,
    "--configdir", config_dir,
]

# Add DAF flag if enabled
if args.DAF:
    command.append("--DAF")

# Add logging level control
if args.debug:
    command.append("--debug")
else:
    command.extend(["--log-level", args.log_level])

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
