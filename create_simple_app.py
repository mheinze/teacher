#!/usr/bin/env python3
"""
Simple macOS Application Bundle Creator
Creates a basic .app bundle structure that can run our Python GUI
"""

import os
import shutil
import stat
import subprocess
import sys

def create_simple_app_bundle():
    """Create a simple .app bundle that launches our GUI"""
    
    app_name = "AIG Class List Processor"
    app_bundle = f"{app_name}.app"
    
    print(f"Creating simple app bundle: {app_bundle}")
    
    # Remove existing bundle
    if os.path.exists(app_bundle):
        shutil.rmtree(app_bundle)
    
    # Create bundle structure
    contents_dir = os.path.join(app_bundle, "Contents")
    macos_dir = os.path.join(contents_dir, "MacOS")
    resources_dir = os.path.join(contents_dir, "Resources")
    
    os.makedirs(macos_dir)
    os.makedirs(resources_dir)
    
    # Create Info.plist
    info_plist = f"""<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
    <key>CFBundleExecutable</key>
    <string>{app_name}</string>
    <key>CFBundleIdentifier</key>
    <string>com.teacher.aig-processor</string>
    <key>CFBundleInfoDictionaryVersion</key>
    <string>6.0</string>
    <key>CFBundleName</key>
    <string>{app_name}</string>
    <key>CFBundlePackageType</key>
    <string>APPL</string>
    <key>CFBundleVersion</key>
    <string>1.0.0</string>
    <key>LSMinimumSystemVersion</key>
    <string>10.13</string>
    <key>NSHighResolutionCapable</key>
    <true/>
    <key>LSUIElement</key>
    <false/>
</dict>
</plist>"""
    
    with open(os.path.join(contents_dir, "Info.plist"), "w") as f:
        f.write(info_plist)
    
    # Create launcher script
    launcher_script = f"""#!/bin/bash
# Launcher script for {app_name}

# Get the directory where this script is located
DIR="$( cd "$( dirname "${{BASH_SOURCE[0]}}" )" && pwd )"
BUNDLE_DIR="$DIR/../.."
RESOURCES_DIR="$BUNDLE_DIR/Contents/Resources"

# Change to the resources directory
cd "$RESOURCES_DIR"

# Find Python executable
PYTHON_EXEC=""
if command -v python3 &> /dev/null; then
    PYTHON_EXEC="python3"
elif command -v python &> /dev/null; then
    PYTHON_EXEC="python"
else
    osascript -e 'display alert "Python Not Found" message "Python 3 is required but not found on this system. Please install Python 3 first." buttons {{"OK"}} default button 1'
    exit 1
fi

# Set up environment
export PYTHONPATH="$RESOURCES_DIR:$PYTHONPATH"

# Launch the GUI with error handling
$PYTHON_EXEC aig_gui.py 2>&1 | logger -t "{app_name}"

# If the Python script failed, show an error
if [ $? -ne 0 ]; then
    osascript -e 'display alert "Launch Error" message "The application failed to start. Please check Console.app for error messages tagged with \\"{app_name}\\"." buttons {{"OK"}} default button 1'
fi
"""
    
    launcher_path = os.path.join(macos_dir, app_name)
    with open(launcher_path, "w") as f:
        f.write(launcher_script)
    
    # Make launcher executable
    st = os.stat(launcher_path)
    os.chmod(launcher_path, st.st_mode | stat.S_IEXEC)
    
    # Copy Python files to Resources
    files_to_copy = [
        "aig_gui.py",
        "aig_processor.py", 
        "requirements.txt"
    ]
    
    for file in files_to_copy:
        if os.path.exists(file):
            shutil.copy2(file, resources_dir)
            print(f"Copied {file} to app bundle")
    
    # Copy input directory if it exists
    if os.path.exists("input"):
        shutil.copytree("input", os.path.join(resources_dir, "input"))
        print("Copied input directory to app bundle")
    
    print(f"✅ Created {app_bundle}")
    print(f"The app will use the system Python installation")
    print(f"Make sure Python 3 and required packages are installed system-wide")
    
    return app_bundle

if __name__ == "__main__":
    create_simple_app_bundle()
