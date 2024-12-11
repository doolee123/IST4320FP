import subprocess
import sys
import os

def install_dependencies():
    """Install required packages for the workout planner application."""
    
    # List of required packages
    requirements = [
        'tk',  # Basic tkinter
        'ttkthemes',  # Themed widgets
        'tkcalendar',  # Calendar widget
        'pillow',  # Image processing
        'matplotlib',  # Plotting
        'openpyxl'  # Excel file handling
    ]

    print("Starting dependency installation...\n")

    # Check if pip is installed
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', '--version'])
    except subprocess.CalledProcessError:
        print("pip is not installed. Please install pip first.")
        return

    # Install each package
    for package in requirements:
        print(f"Installing {package}...")
        try:
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
            print(f"Successfully installed {package}")
        except subprocess.CalledProcessError as e:
            print(f"Failed to install {package}. Error: {str(e)}")
        print()

    print("\nAll dependencies have been processed.")
    print("\nIf you encounter any issues, try running:")
    print("pip install tk ttkthemes tkcalendar pillow matplotlib openpyxl")

if __name__ == "__main__":
    install_dependencies()