import subprocess
import sys

def check_and_install(package):
    try:
        __import__(package)
        print(f"{package} is already installed.")
    except ImportError:
        print(f"{package} is not installed. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])

if __name__ == "__main__":
    required_packages = ["xlwings", "xlsxwriter"]
    for package in required_packages:
        check_and_install(package)