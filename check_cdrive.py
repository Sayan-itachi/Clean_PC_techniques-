#!/usr/bin/env python3
"""
Windows C: Drive Storage Audit Script
Performs two-pass analysis to identify large files and categorize them for cleanup
"""

import os
import sys
import pandas as pd
from pathlib import Path
from collections import defaultdict
import time

def bytes_to_gb(bytes_size):
    """Convert bytes to GB with 2 decimal places"""
    return round(bytes_size / (1024**3), 2)

def get_folder_size(folder_path, max_depth=None, current_depth=0):
    """
    Calculate total size of a folder
    If max_depth is specified, only go that deep
    """
    total_size = 0
    try:
        if max_depth is not None and current_depth > max_depth:
            return 0
            
        for item in os.listdir(folder_path):
            item_path = os.path.join(folder_path, item)
            try:
                if os.path.isfile(item_path):
                    total_size += os.path.getsize(item_path)
                elif os.path.isdir(item_path):
                    total_size += get_folder_size(item_path, max_depth, current_depth + 1)
            except (OSError, PermissionError):
                continue
    except (OSError, PermissionError):
        pass
    
    return total_size

def categorize_file(file_path):
    """Categorize files based on path and extension"""
    file_path_lower = file_path.lower()
    file_ext = Path(file_path).suffix.lower()
    
    # System Critical - Windows system folders
    if '\\windows\\system32\\' in file_path_lower or '\\windows\\winsxs\\' in file_path_lower:
        return 'SYSTEM_CRITICAL'
    
    # Likely Useless extensions and paths
    useless_extensions = {'.exe', '.zip', '.iso', '.msi', '.rar', '.7z', '.bak', '.tmp', '.log'}
    useless_paths = ['\\downloads\\', '\\temp\\', '\\cache\\', '\\appdata\\local\\temp\\']
    
    if file_ext in useless_extensions or any(path in file_path_lower for path in useless_paths):
        return 'LIKELY_USELESS'
    
    # Media files
    media_extensions = {'.mp4', '.mkv', '.avi', '.mov', '.mp3', '.wav', '.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webm', '.flv', '.wmv'}
    if file_ext in media_extensions:
        return 'MEDIA'
    
    # Everything else
    return 'REVIEW'

def scan_large_files(folder_path, min_size_mb=50, max_depth=3):
    """
    Scan for files larger than min_size_mb in a folder up to max_depth
    Returns list of (file_path, size_bytes, category)
    """
    large_files = []
    min_size_bytes = min_size_mb * 1024 * 1024
    
    def scan_recursive(current_path, current_depth=0):
        if current_depth > max_depth:
            return
            
        try:
            for item in os.listdir(current_path):
                item_path = os.path.join(current_path, item)
                try:
                    if os.path.isfile(item_path):
                        file_size = os.path.getsize(item_path)
                        if file_size >= min_size_bytes:
                            category = categorize_file(item_path)
                            large_files.append((item_path, file_size, category))
                    elif os.path.isdir(item_path):
                        scan_recursive(item_path, current_depth + 1)
                except (OSError, PermissionError):
                    continue
        except (OSError, PermissionError):
            pass
    
    scan_recursive(folder_path)
    return large_files

def main():
    print("Windows C: Drive Storage Audit")
    print("=" * 50)
    print()
    
    # Check if running on Windows
    if os.name != 'nt':
        print("This script is designed to run on Windows.")
        sys.exit(1)
    
    c_drive = "C:\\"
    if not os.path.exists(c_drive):
        print("C:\\ drive not found.")
        sys.exit(1)
    
    # PASS 1: Scan top-level folders
    print("=== PASS 1: Largest Top-Level Folders (Depth=1) ===")
    
    folder_sizes = []
    total_scanned = 0
    
    try:
        for item in os.listdir(c_drive):
            item_path = os.path.join(c_drive, item)
            if os.path.isdir(item_path):
                print(f"Scanning {item_path}...", end=' ')
                folder_size = get_folder_size(item_path, max_depth=10)  # Reasonable depth for top-level scan
                folder_size_gb = bytes_to_gb(folder_size)
                folder_sizes.append((item_path, folder_size_gb))
                total_scanned += folder_size
                print(f"{folder_size_gb:.2f} GB")
    except (OSError, PermissionError) as e:
        print(f"Error scanning C:\\: {e}")
        sys.exit(1)
    
    # Sort by size (descending)
    folder_sizes.sort(key=lambda x: x[1], reverse=True)
    
    # Display table
    print()
    print(f"{'Folder':<40} {'Size_GB':>10}")
    print("-" * 52)
    for folder_path, size_gb in folder_sizes:
        print(f"{folder_path:<40} {size_gb:>10.2f}")
    print("-" * 52)
    print(f"{'Total scanned:':<40} {bytes_to_gb(total_scanned):>10.2f} GB")
    print()
    
    # PASS 2: Detailed scan of top 3 folders
    print("=== PASS 2: Large Files in Top 3 Folders (50+ MB, Depth=3) ===")
    print()
    
    top_3_folders = folder_sizes[:3]
    all_large_files = []
    category_totals = defaultdict(float)
    
    for folder_path, _ in top_3_folders:
        print(f"Scanning large files in: {folder_path}")
        large_files = scan_large_files(folder_path, min_size_mb=50, max_depth=3)
        
        if not large_files:
            print(f"No files ≥50 MB found in {folder_path}")
            print()
            continue
        
        # Group by category
        files_by_category = defaultdict(list)
        for file_path, file_size, category in large_files:
            files_by_category[category].append((file_path, file_size))
            all_large_files.append((file_path, bytes_to_gb(file_size), category))
            category_totals[category] += bytes_to_gb(file_size)
        
        # Display results for this folder
        print(f"\nFolder: {folder_path}")
        
        # Define category display order and descriptions
        category_descriptions = {
            'LIKELY_USELESS': 'LIKELY_USELESS (Safe to Delete Manually)',
            'MEDIA': 'MEDIA (Review Before Deleting)',
            'REVIEW': 'REVIEW (Manual Review Required)',
            'SYSTEM_CRITICAL': 'SYSTEM_CRITICAL (Do Not Delete)'
        }
        
        for category in ['LIKELY_USELESS', 'MEDIA', 'REVIEW', 'SYSTEM_CRITICAL']:
            if category in files_by_category:
                print(f"\nCategory: {category_descriptions[category]}")
                # Sort files by size (descending)
                files_by_category[category].sort(key=lambda x: x[1], reverse=True)
                for file_path, file_size in files_by_category[category]:
                    print(f"{bytes_to_gb(file_size):>5.2f} GB   {file_path}")
        
        print()
    
    # Summary by category
    print("=== SUMMARY BY CATEGORY ===")
    for category in ['LIKELY_USELESS', 'MEDIA', 'REVIEW', 'SYSTEM_CRITICAL']:
        if category in category_totals:
            print(f"{category:<18} : {category_totals[category]:>7.2f} GB")
    
    print()
    print("Next Steps:")
    print("    1. Start by deleting LIKELY_USELESS files listed above.")
    print("    2. Review MEDIA files — delete only if backed up.")
    print("    3. Avoid anything in SYSTEM_CRITICAL.")
    print()
    
    # Export to CSV
    if all_large_files:
        csv_filename = "C_Drive_Audit_Report.csv"
        df = pd.DataFrame(all_large_files, columns=['Path', 'Size_GB', 'Category'])
        df = df.sort_values(['Category', 'Size_GB'], ascending=[True, False])
        df.to_csv(csv_filename, index=False)
        print(f"Full report exported to: {csv_filename}")
        print(f"Total files in report: {len(all_large_files)}")
    else:
        print("No large files found to export.")
    
    print("\n" + "=" * 50)
    print("Audit completed successfully!")

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nAudit interrupted by user.")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)