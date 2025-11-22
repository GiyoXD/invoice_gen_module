import subprocess
import json
import os
from pathlib import Path
import sys

def run_invoice_generation_job(input_file):
    print(f"--- [Backend] Received request for {input_file} ---")
    
    # 1. Run the script (simulating a background worker)
    # In a real app, this might be Celery or a subprocess
    output_file = "result_demo.xlsx"
    cmd = [sys.executable, "invoice_generator/generate_invoice.py", input_file, "--output", output_file]
    
    print(f"--- [Backend] Launching subprocess: {' '.join(cmd)} ---")
    # We suppress stdout/stderr to show how we rely purely on metadata
    subprocess.run(cmd, capture_output=True)
    
    # 2. Check for Metadata File
    # We know the output is result_demo.xlsx, so metadata is result_demo.xlsx.meta.json
    meta_path = Path(output_file + ".meta.json")
    
    if not meta_path.exists():
        print("❌ [Backend] Error: Metadata file not found! The script might have crashed hard.")
        return

    # 3. Read Metadata
    try:
        with open(meta_path, "r") as f:
            meta = json.load(f)
    except Exception as e:
        print(f"❌ [Backend] Error reading metadata: {e}")
        return
        
    # 4. Use Metadata to update UI
    print("\n--- [Backend] Job Finished. Simulating UI Update... ---")
    
    if meta["status"] == "success":
        print(f"✅ [UI] Success! Invoice generated in {meta['execution_time']:.2f}s")
        print(f"📄 [UI] Download ready at: {meta['output_file']}")
        print(f"📊 [UI] Processed sheets: {', '.join(meta['sheets_processed'])}")
        
        # Example: Check for warnings/partial failures
        if meta.get("sheets_failed"):
             print(f"⚠️ [UI] Warning: Some sheets failed: {', '.join(meta['sheets_failed'])}")
             
    else:
        print(f"⚠️ [UI] Failed! Status: {meta['status']}")
        print(f"🛑 [UI] Error Message: {meta.get('error_message')}")
        
    print("--------------------------------------------------")

if __name__ == "__main__":
    # Simulate a user uploading a file
    # Ensure we are in the right directory
    if not Path("invoice_generator/JF.json").exists():
        print("Error: Please run this script from the root 'refactor' directory.")
    else:
        run_invoice_generation_job("invoice_generator/JF.json")
