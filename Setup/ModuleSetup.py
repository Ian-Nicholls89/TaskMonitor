import importlib
import subprocess
import sys
import os

# Set console to use UTF-8 encoding
if sys.platform == 'win32':
    import ctypes
    kernel32 = ctypes.windll.kernel32
    kernel32.SetConsoleCP(65001)
    kernel32.SetConsoleOutputCP(65001)

# ASCII alternatives for status indicators
STATUS_OK = "[OK]"
STATUS_INSTALL = "[INSTALL]"
STATUS_ERROR = "[ERROR]"
STATUS_WARN = "[WARN]"

try:
    print("Module setup started...\n")
    print(f"Current working directory: {os.getcwd()}")
    
    # Try to find and read requirements.txt
    requirements_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'requirements.txt')
    print(f"Looking for requirements.txt at: {requirements_path}")
    
    if not os.path.exists(requirements_path):
        print(f"{STATUS_ERROR} requirements.txt not found at {requirements_path}")
        print("Please make sure the file exists in the parent directory.")
        sys.exit(1)
    
    with open(requirements_path, "r") as f:
        modules = [line.strip() for line in f if line.strip() and not line.startswith('#')]
    
    print(f"Found {len(modules)} modules to check/install")
    
    if not modules:
        print(f"{STATUS_WARN} No modules found in requirements.txt")
    else:
        print("\nChecking and installing required Python modules...")
        for module in modules:
            try:
                print(f"\nChecking: {module}")
                importlib.import_module(module)
                print(f"{STATUS_OK} {module} is already available.")
            except ImportError:
                print(f"{STATUS_INSTALL} {module} not found. Installing...")
                try:
                    subprocess.check_call(
                        [sys.executable, "-m", "pip", "install", module],
                        stdout=subprocess.DEVNULL,
                        stderr=subprocess.STDOUT
                    )
                    print(f"{STATUS_OK} Successfully installed {module}.")
                except subprocess.CalledProcessError as e:
                    print(f"{STATUS_ERROR} Failed to install {module}. Error code: {e.returncode}")
                    if hasattr(e, 'output') and e.output:
                        try:
                            print(f"Error output: {e.output.decode('utf-8', errors='replace')}")
                        except:
                            print("(Could not decode error output)")
                except Exception as e:
                    print(f"{STATUS_ERROR} Unexpected error installing {module}: {str(e)}")
    
    print(f"\n{STATUS_OK} Module setup completed!")
    
except Exception as e:
    print(f"\n{STATUS_ERROR} An unexpected error occurred: {str(e)}", file=sys.stderr)
    import traceback
    traceback.print_exc()
    sys.exit(1)
