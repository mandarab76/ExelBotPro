#!/usr/bin/env python3
"""
ExcelBot Pro Package Validation Script
This script validates the complete Office Add-in package structure
"""

import os
import json
import xml.etree.ElementTree as ET
from pathlib import Path


def check_file_exists(filepath, description):
    """Check if a file exists and print result"""
    if os.path.exists(filepath):
        size = os.path.getsize(filepath)
        print(f"✓ {description}: {filepath} ({size} bytes)")
        return True
    else:
        print(f"✗ {description} missing: {filepath}")
        return False


def validate_manifest():
    """Validate the Office Add-in manifest.xml"""
    print("\n=== Validating manifest.xml ===")
    try:
        tree = ET.parse('manifest.xml')
        root = tree.getroot()
        print("✓ manifest.xml is valid XML")
        
        # Check for required elements
        required_elements = ['Id', 'Version', 'ProviderName', 'DefaultLocale', 
                           'DisplayName', 'Description', 'Hosts', 'DefaultSettings']
        
        for elem in required_elements:
            found = root.find('.//{http://schemas.microsoft.com/office/appforoffice/1.1}' + elem)
            if found is not None:
                print(f"✓ Required element: {elem}")
            else:
                print(f"✗ Missing required element: {elem}")
                return False
        
        return True
    except Exception as e:
        print(f"✗ Error validating manifest: {e}")
        return False


def validate_package_json():
    """Validate package.json structure"""
    print("\n=== Validating package.json ===")
    try:
        with open('package.json', 'r') as f:
            pkg = json.load(f)
        
        print("✓ package.json is valid JSON")
        
        required_fields = ['name', 'version', 'description', 'scripts']
        for field in required_fields:
            if field in pkg:
                value = pkg.get(field)
                if field == 'scripts':
                    print(f"✓ Has {field}: {len(value)} scripts")
                else:
                    print(f"✓ Has {field}: {value}")
            else:
                print(f"✗ Missing {field}")
                return False
        
        # Check for important scripts
        required_scripts = ['start', 'build', 'validate']
        for script in required_scripts:
            if script in pkg.get('scripts', {}):
                print(f"✓ Script defined: {script}")
            else:
                print(f"⚠ Optional script missing: {script}")
        
        return True
    except Exception as e:
        print(f"✗ Error validating package.json: {e}")
        return False


def validate_html():
    """Validate HTML files"""
    print("\n=== Validating HTML Files ===")
    try:
        with open('src/taskpane/taskpane.html', 'r') as f:
            html = f.read()
        
        checks = [
            ('<!DOCTYPE html>', 'DOCTYPE declaration'),
            ('<html>', 'HTML tag'),
            ('<head>', 'Head section'),
            ('<body>', 'Body section'),
            ('office.js', 'Office.js reference'),
            ('taskpane.css', 'CSS reference'),
            ('taskpane.js', 'JavaScript reference'),
        ]
        
        all_passed = True
        for check, description in checks:
            if check in html:
                print(f"✓ {description}")
            else:
                print(f"✗ Missing {description}")
                all_passed = False
        
        return all_passed
    except Exception as e:
        print(f"✗ Error validating HTML: {e}")
        return False


def validate_javascript():
    """Validate JavaScript files"""
    print("\n=== Validating JavaScript Files ===")
    try:
        with open('src/taskpane/taskpane.js', 'r') as f:
            js = f.read()
        
        checks = [
            ('Office.onReady', 'Office.js initialization'),
            ('Office.HostType.Excel', 'Excel host check'),
            ('async function', 'Async functions'),
            ('Excel.run', 'Excel API usage'),
            ('addEventListener', 'Event handlers'),
        ]
        
        for check, description in checks:
            if check in js:
                print(f"✓ {description}")
            else:
                print(f"⚠ {description} not found (might be optional)")
        
        print(f"✓ JavaScript file size: {len(js)} bytes")
        return True
    except Exception as e:
        print(f"✗ Error validating JavaScript: {e}")
        return False


def validate_vba():
    """Validate VBA module files"""
    print("\n=== Validating VBA Modules ===")
    vba_files = [
        'src/vba/ExcelBotProCore.bas',
        'src/vba/ExcelBotProUtils.bas'
    ]
    
    all_valid = True
    for vba_file in vba_files:
        try:
            with open(vba_file, 'r') as f:
                content = f.read()
            
            print(f"\n{vba_file}:")
            
            checks = [
                ('Attribute VB_Name', 'Module name'),
                ('Option Explicit', 'Option Explicit'),
                ('Sub ', 'Subroutines'),
                ('End Sub', 'End Sub statements'),
            ]
            
            for check, description in checks:
                if check in content:
                    print(f"  ✓ {description}")
                else:
                    print(f"  ✗ Missing {description}")
                    all_valid = False
            
            # Count procedures
            sub_count = content.count('Sub ')
            func_count = content.count('Function ')
            print(f"  ✓ Contains {sub_count} Subs and {func_count} Functions")
            
        except Exception as e:
            print(f"✗ Error validating {vba_file}: {e}")
            all_valid = False
    
    return all_valid


def validate_directory_structure():
    """Validate the directory structure"""
    print("\n=== Validating Directory Structure ===")
    
    required_files = [
        ('manifest.xml', 'Office Add-in Manifest'),
        ('package.json', 'Node.js Package File'),
        ('webpack.config.js', 'Webpack Configuration'),
        ('.gitignore', 'Git Ignore File'),
        ('README.md', 'Documentation'),
        ('INSTALL.md', 'Installation Guide'),
        ('src/taskpane/taskpane.html', 'Task Pane HTML'),
        ('src/taskpane/taskpane.css', 'Task Pane CSS'),
        ('src/taskpane/taskpane.js', 'Task Pane JavaScript'),
        ('src/vba/ExcelBotProCore.bas', 'Core VBA Module'),
        ('src/vba/ExcelBotProUtils.bas', 'Utility VBA Module'),
    ]
    
    all_exist = True
    for filepath, description in required_files:
        if not check_file_exists(filepath, description):
            all_exist = False
    
    return all_exist


def main():
    """Run all validations"""
    print("=" * 60)
    print("ExcelBot Pro Package Validation")
    print("=" * 60)
    
    os.chdir(Path(__file__).parent)
    
    results = {
        'Directory Structure': validate_directory_structure(),
        'Manifest': validate_manifest(),
        'Package.json': validate_package_json(),
        'HTML': validate_html(),
        'JavaScript': validate_javascript(),
        'VBA Modules': validate_vba(),
    }
    
    print("\n" + "=" * 60)
    print("Validation Summary")
    print("=" * 60)
    
    all_passed = True
    for category, passed in results.items():
        status = "✓ PASS" if passed else "✗ FAIL"
        print(f"{status:8} {category}")
        if not passed:
            all_passed = False
    
    print("=" * 60)
    
    if all_passed:
        print("\n🎉 All validations passed! Package is ready.")
        return 0
    else:
        print("\n⚠️  Some validations failed. Please review the output above.")
        return 1


if __name__ == '__main__':
    exit(main())
