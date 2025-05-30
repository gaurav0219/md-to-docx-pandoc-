#!/usr/bin/env python3
"""
Build script for creating executable from the MD to DOCX Converter
"""

import subprocess
import sys
import os
import shutil
from pathlib import Path

def main():
    """Main build function"""
    print("=" * 60)
    print("MD to DOCX Converter - Executable Builder")
    print("=" * 60)
    
    # Clean previous builds
    clean_previous_builds()
    
    # Build the executable
    build_executable()
    
    print("\nBuild process completed!")
    print("=" * 60)

def clean_previous_builds():
    """Clean previous build artifacts"""
    print("\n[1/3] Cleaning previous builds...")
    
    # Directories to clean
    dirs_to_clean = ['build', 'dist', '__pycache__']
    
    for dir_name in dirs_to_clean:
        dir_path = Path(dir_name)
        if dir_path.exists():
            print(f"  Removing {dir_path}...")
            try:
                shutil.rmtree(dir_path)
                print(f"  ✓ {dir_path} removed successfully")
            except Exception as e:
                print(f"  ✗ Failed to remove {dir_path}: {e}")
    
    # Remove spec file if it exists
    spec_file = Path('md_to_docx_converter.spec')
    if spec_file.exists():
        try:
            spec_file.unlink()
            print(f"  ✓ {spec_file} removed successfully")
        except Exception as e:
            print(f"  ✗ Failed to remove {spec_file}: {e}")

def build_executable():
    """Build the executable using PyInstaller"""
    print("\n[2/3] Building executable...")
    
    # Ensure icon exists
    icon_path = Path('icon.png')
    if not icon_path.exists():
        print(f"  ! Warning: Icon file {icon_path} not found")
        icon_arg = []
    else:
        icon_arg = ['--icon', str(icon_path)]
    
    # Build command
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name', 'MD_to_DOCX_Converter',
        '--onefile',
        '--windowed',
        *icon_arg,
        '--add-data', f"{icon_path};.",
        'main.py'
    ]
    
    # On non-Windows platforms, use : instead of ; for path separator
    if os.name != 'nt':
        cmd[10] = f"{icon_path}:."
    
    print(f"  Running: {' '.join(cmd)}")
    
    try:
        subprocess.run(cmd, check=True)
        print("  ✓ Executable built successfully")
        
        # Check if build was successful
        output_file = Path('dist/MD_to_DOCX_Converter.exe') if os.name == 'nt' else Path('dist/MD_to_DOCX_Converter')
        if output_file.exists():
            print(f"\n[3/3] Success! Executable created at: {output_file.absolute()}")
        else:
            print(f"\n[3/3] ✗ Failed to find output executable at expected location: {output_file.absolute()}")
    except subprocess.CalledProcessError as e:
        print(f"  ✗ Failed to build executable: {e}")
        print(f"\n[3/3] ✗ Build process failed")

if __name__ == "__main__":
    main()
