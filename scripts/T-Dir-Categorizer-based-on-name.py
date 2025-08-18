#!/usr/bin/env python3
"""
File Categorization Script
Automatically categorizes files into folders based on filename patterns.
Creates copies for files that match multiple categories.
"""

import os
import shutil
import re
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from collections import defaultdict

class FileCategorizer:
    def __init__(self):
        # Define categorization rules - add/modify patterns as needed
        self.categories = {
            'T1D': [r't1d', r'type.*1.*diabetes', r'diabetes.*type.*1'],
            'T2D': [r't2d', r'type.*2.*diabetes', r'diabetes.*type.*2'],
            'Control': [r'control', r'ctrl', r'normal', r'healthy'],
            'Big_Duct': [r'big.*duct', r'large.*duct', r'main.*duct', r'major.*duct'],
            'Small_Duct': [r'small.*duct', r'minor.*duct', r'micro.*duct'],
            'Islet': [r'islet', r'pancreatic.*islet', r'langerhans'],
            'High_Magnification': [r'high.*mag', r'20x', r'40x', r'60x', r'100x', r'zoom'],
            'Low_Magnification': [r'low.*mag', r'4x', r'5x', r'10x', r'overview'],
            'Positive': [r'positive', r'pos', r'\+', r'stain.*pos'],
            'Negative': [r'negative', r'neg', r'\-', r'stain.*neg'],
            'miR155': [r'mir.*155', r'mir-155', r'microRNA.*155'],
            'H_E': [r'h&e', r'he', r'hematoxylin.*eosin', r'h_e'],
            'DAB': [r'dab', r'diaminobenzidine', r'brown.*stain'],
            'DAPI': [r'dapi', r'nuclei.*stain', r'blue.*nuclei']
        }
        
        self.source_dir = None
        self.categorized_files = defaultdict(list)
        self.file_matches = defaultdict(list)  # Track which categories each file matches
        
    def choose_directory(self):
        """Let user choose source directory using tkinter"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        self.source_dir = filedialog.askdirectory(
            title="Select directory containing files to categorize"
        )
        
        if not self.source_dir:
            messagebox.showinfo("Cancelled", "No directory selected. Exiting.")
            return False
            
        print(f"Selected directory: {self.source_dir}")
        return True
    
    def get_file_categories(self, filename):
        """Determine which categories a file belongs to based on its name"""
        filename_lower = filename.lower()
        matches = []
        
        for category, patterns in self.categories.items():
            for pattern in patterns:
                if re.search(pattern, filename_lower, re.IGNORECASE):
                    matches.append(category)
                    break  # Found a match for this category, move to next
                    
        return matches
    
    def scan_files(self):
        """Scan directory and categorize all files"""
        if not self.source_dir:
            print("No source directory selected!")
            return
            
        print("Scanning files...")
        total_files = 0
        categorized_files = 0
        
        for file_path in Path(self.source_dir).iterdir():
            if file_path.is_file():
                total_files += 1
                filename = file_path.name
                categories = self.get_file_categories(filename)
                
                if categories:
                    categorized_files += 1
                    self.file_matches[filename] = categories
                    
                    for category in categories:
                        self.categorized_files[category].append(file_path)
                        
                    # Show multiple matches
                    if len(categories) > 1:
                        print(f"📁 {filename} → {', '.join(categories)} (MULTIPLE)")
                    else:
                        print(f"📄 {filename} → {categories[0]}")
                else:
                    print(f"❓ {filename} → No category match")
        
        print(f"\nScan complete:")
        print(f"Total files: {total_files}")
        print(f"Categorized files: {categorized_files}")
        print(f"Uncategorized files: {total_files - categorized_files}")
        
        # Show category summary
        print(f"\nCategory breakdown:")
        for category, files in self.categorized_files.items():
            print(f"  {category}: {len(files)} files")
            
        return total_files > 0
    
    def create_category_folders(self):
        """Create folders for each category that has files"""
        categories_dir = Path(self.source_dir) / "Categorized"
        categories_dir.mkdir(exist_ok=True)
        
        for category in self.categorized_files.keys():
            category_path = categories_dir / category
            category_path.mkdir(exist_ok=True)
            print(f"Created folder: {category_path}")
            
        return categories_dir
    
    def copy_files_to_categories(self):
        """Copy files to their respective category folders"""
        if not self.categorized_files:
            print("No files to categorize!")
            return
            
        categories_dir = self.create_category_folders()
        copy_count = 0
        duplicate_count = 0
        
        print("\nCopying files to category folders...")
        
        for category, file_paths in self.categorized_files.items():
            category_folder = categories_dir / category
            
            for file_path in file_paths:
                dest_path = category_folder / file_path.name
                
                # Handle filename conflicts within the same category
                counter = 1
                original_dest = dest_path
                while dest_path.exists():
                    stem = original_dest.stem
                    suffix = original_dest.suffix
                    dest_path = category_folder / f"{stem}_copy{counter}{suffix}"
                    counter += 1
                
                try:
                    shutil.copy2(file_path, dest_path)
                    copy_count += 1
                    
                    # Check if this is a duplicate (file appears in multiple categories)
                    if len(self.file_matches[file_path.name]) > 1:
                        duplicate_count += 1
                        
                except Exception as e:
                    print(f"❌ Error copying {file_path.name} to {category}: {e}")
        
        print(f"\nCopy complete:")
        print(f"Total file copies created: {copy_count}")
        print(f"Files duplicated across categories: {duplicate_count}")
        
        # Show duplicate summary
        duplicates = {f: cats for f, cats in self.file_matches.items() if len(cats) > 1}
        if duplicates:
            print(f"\nFiles copied to multiple folders:")
            for filename, categories in duplicates.items():
                print(f"  📋 {filename} → {', '.join(categories)}")
    
    def add_custom_patterns(self):
        """Allow user to add custom categorization patterns"""
        root = tk.Tk()
        root.withdraw()
        
        response = messagebox.askyesno(
            "Custom Patterns", 
            "Would you like to add custom categorization patterns?"
        )
        
        if response:
            print("\nAdd custom patterns (format: CategoryName:pattern1,pattern2,pattern3)")
            print("Example: MyCategory:experiment,test,sample")
            print("Enter empty line to finish:")
            
            while True:
                user_input = input("Enter pattern: ").strip()
                if not user_input:
                    break
                    
                try:
                    category, patterns_str = user_input.split(':', 1)
                    patterns = [p.strip() for p in patterns_str.split(',')]
                    self.categories[category.strip()] = patterns
                    print(f"Added category '{category}' with patterns: {patterns}")
                except ValueError:
                    print("Invalid format. Use: CategoryName:pattern1,pattern2")
    
    def run(self):
        """Main execution method"""
        print("🗂️  File Categorization Script")
        print("=" * 50)
        
        # Choose directory
        if not self.choose_directory():
            return
            
        # Option to add custom patterns
        self.add_custom_patterns()
        
        # Scan and categorize files
        if not self.scan_files():
            print("No files found in the selected directory!")
            return
            
        # Confirm before copying
        root = tk.Tk()
        root.withdraw()
        
        proceed = messagebox.askyesno(
            "Confirm Copy",
            f"Ready to copy {sum(len(files) for files in self.categorized_files.values())} "
            f"file instances to {len(self.categorized_files)} category folders.\n\n"
            "This will create a 'Categorized' folder in your selected directory.\n\n"
            "Proceed?"
        )
        
        if proceed:
            self.copy_files_to_categories()
            
            messagebox.showinfo(
                "Complete", 
                f"Categorization complete!\n\n"
                f"Check the 'Categorized' folder in:\n{self.source_dir}"
            )
        else:
            print("Operation cancelled by user.")

def main():
    """Main function"""
    try:
        categorizer = FileCategorizer()
        categorizer.run()
    except KeyboardInterrupt:
        print("\n\nOperation cancelled by user.")
    except Exception as e:
        print(f"\nError: {e}")
        messagebox.showerror("Error", f"An error occurred: {e}")

if __name__ == "__main__":
    main()
