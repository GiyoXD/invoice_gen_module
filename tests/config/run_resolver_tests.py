"""
Quick test runner for TableDataResolver

Usage:
    python tests/config/run_resolver_tests.py
    python tests/config/run_resolver_tests.py --verbose
    python tests/config/run_resolver_tests.py --coverage
"""
import sys
import subprocess
from pathlib import Path

def run_tests(verbose=False, coverage=False):
    """Run TableDataResolver tests."""
    test_file = Path(__file__).parent / "test_table_data_resolver.py"
    
    cmd = ["python", "-m", "pytest", str(test_file)]
    
    if verbose:
        cmd.extend(["-v", "-s"])
    
    if coverage:
        cmd.extend([
            "--cov=invoice_generator.config.table_data_resolver",
            "--cov-report=term-missing",
            "--cov-report=html"
        ])
    
    print(f"Running: {' '.join(cmd)}\n")
    result = subprocess.run(cmd, cwd=Path(__file__).parent.parent.parent)
    return result.returncode

if __name__ == "__main__":
    verbose = "--verbose" in sys.argv or "-v" in sys.argv
    coverage = "--coverage" in sys.argv or "--cov" in sys.argv
    
    exit_code = run_tests(verbose=verbose, coverage=coverage)
    sys.exit(exit_code)
