import logging

logging.basicConfig(format='%(asctime)s.%(msecs)03d %(levelname)-8s [%(filename)s:%(lineno)d] %(message)s',
        datefmt='%H:%M:%S', level=logging.DEBUG)
basic_logger = logging.getLogger('BASIC')
basic_logger.info(f'start {__file__}')

import sys
import subprocess

# cargo build artifacts on the D: drive
test_target_dir = r"D:/Elements_only/ephem/apps/rust/auto_generated/undoc_testing"

extra_args = sys.argv[1:]

command = [
    "uv", "run", "cargo", "test",
    "--target-dir", test_target_dir
] + extra_args

print("=" * 56)
print(f"RUNNING UNDOC TESTS: {' '.join(command)}")
print("=" * 56)

subprocess.run(command)