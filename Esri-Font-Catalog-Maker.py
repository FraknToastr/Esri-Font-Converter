"""
ESRI Font Catalog Generator with Character Mappings
Creates a PDF displaying all ESRI fonts with complete character grids
showing symbols and their Unicode/ASCII mappings
"""

import sys
import subprocess
import os
from datetime import datetime

# Configuration
output_pdf = r"C:\Temp\ESRI_Font_Character_Catalog.pdf"

print("=" * 70)
print("ESRI Font Character Catalog Generator")
print("=" * 70)
print(f"Python: {sys.version}")
print(f"Environment: {sys.executable}\n")

# Check and install required packages
required_packages = {
    'reportlab': 'reportlab',
}

def check_and_install_package(package_name, pip_name):
    """Check if package is installed, install if not"""
    try:
        __import__(package_name)
        print(f"✓ {package_name} is already installed")
        return True
    except ImportError:
        print(f"✗ {package_name} is not installed")
        print(f"  Installing {pip_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", pip_name])
            print(f"✓ Successfully installed {pip_name}")
            return True
        except subprocess.CalledProcessError as e:
            print(f"✗ Failed to install {pip_name}: {e}")
            return False

print("Checking dependencies...")
print("-" * 70)

all_installed = True
for package_name, pip_name in required_packages.items():
    if not check_and_install_package(package_name, pip_name):
        all_installed = False

if not all_installed:
    print("\n" + "=" * 70)
    print("ERROR: Could not install all required packages")
    print("Please install manually using: pip install reportlab")
    print("=" * 70)
    sys.exit(1)

print("-" * 70)
print("All dependencies satisfied!\n")

# Now import the packages
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Create output directory if it doesn't exist
output_dir = os.path.dirname(output_pdf)
if not os.path.exists(output_dir):
    os.makedirs(output_dir)

print("=" * 70)
print("SCANNING FOR ESRI FONTS")
print("=" * 70)

# Get ESRI fonts only
font_dict = {}

# Scan ArcGIS Pro font directories
arcgis_font_dirs = [
    r"C:\Program Files\ArcGIS\Pro\Resources\Fonts",
    r"C:\Program Files (x86)\ArcGIS\Desktop10.8\Fonts",  # Desktop fallback
    os.path.join(os.environ.get('PROGRAMFILES', 'C:\\Program Files'), 'ArcGIS', 'Pro', 'Resources', 'Fonts'),
]

fonts_found = False
for font_dir in arcgis_font_dirs:
    if os.path.exists(font_dir):
        print(f"Found directory: {font_dir}")
        fonts_found = True
        for root, dirs, files in os.walk(font_dir):
            for file in files:
                if file.lower().endswith(('.ttf', '.otf')):
                    font_path = os.path.join(root, file)
                    font_name = os.path.splitext(file)[0]
                    if font_name not in font_dict:
                        font_dict[font_name] = font_path
                        print(f"  Added: {font_name}")

if not fonts_found:
    print("\nWARNING: No ArcGIS font directories found!")
    print("Please update the arcgis_font_dirs list with your ArcGIS installation path.")

fonts = sorted(font_dict.keys())
total_fonts = len(fonts)
print(f"\nTotal ESRI fonts found: {total_fonts}\n")

if total_fonts == 0:
    print("ERROR: No fonts found. Exiting.")
    sys.exit(1)

print("=" * 70)
print("GENERATING CHARACTER CATALOG")
print("=" * 70)

# PDF Setup - using landscape for more space
page_width, page_height = landscape(letter)
c = canvas.Canvas(output_pdf, pagesize=landscape(letter))

margin = 0.3 * inch
cell_size = 0.50 * inch
font_size_char = 20
font_size_label = 6
cols = 16  # 16 columns for character grid
rows = 14  # 14 rows for character grid (fits better on page)

def draw_page_header(canvas_obj, font_name, page_num, total_pages):
    """Draw page header with font name"""
    canvas_obj.setFont("Helvetica-Bold", 14)
    canvas_obj.drawString(margin, page_height - 0.3 * inch, f"Font: {font_name}")
    
    canvas_obj.setFont("Helvetica", 8)
    canvas_obj.drawString(page_width - 2 * inch, page_height - 0.3 * inch, 
                         f"Page {page_num} of {total_pages}")
    
    canvas_obj.drawString(margin, page_height - 0.5 * inch,
                         f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    
    # Draw line separator
    canvas_obj.line(margin, page_height - 0.6 * inch, 
                   page_width - margin, page_height - 0.6 * inch)

def draw_character_grid(canvas_obj, font_id, font_name, start_code=32, end_code=255):
    """Draw a grid of characters with their codes"""
    
    y_start = page_height - 0.9 * inch
    x_start = margin
    
    char_count = 0
    
    # Special character names for non-printable/special keys
    special_chars = {
        32: "Space", 33: "!", 34: '"', 35: "#", 36: "$", 37: "%", 38: "&", 39: "'",
        40: "(", 41: ")", 42: "*", 43: "+", 44: ",", 45: "-", 46: ".", 47: "/",
        58: ":", 59: ";", 60: "<", 61: "=", 62: ">", 63: "?", 64: "@",
        91: "[", 92: "\\", 93: "]", 94: "^", 95: "_", 96: "`",
        123: "{", 124: "|", 125: "}", 126: "~", 127: "DEL"
    }
    
    for code in range(start_code, min(end_code + 1, start_code + (rows * cols))):
        row = (code - start_code) // cols
        col = (code - start_code) % cols
        
        if row >= rows:
            break
        
        x = x_start + (col * cell_size)
        y = y_start - (row * cell_size)
        
        # Draw cell border
        canvas_obj.setStrokeColorRGB(0.8, 0.8, 0.8)
        canvas_obj.rect(x, y - cell_size, cell_size, cell_size)
        
        # Keyboard key label (top)
        if code in special_chars:
            key_label = special_chars[code]
        elif 48 <= code <= 57:  # Numbers 0-9
            key_label = chr(code)
        elif 65 <= code <= 90:  # Uppercase A-Z
            key_label = chr(code)
        elif 97 <= code <= 122:  # Lowercase a-z
            key_label = chr(code)
        else:
            key_label = chr(code) if code < 127 else f"U+{code:04X}"
        
        canvas_obj.setFont("Helvetica-Bold", font_size_label)
        canvas_obj.setFillColorRGB(0.2, 0.2, 0.2)
        # Truncate long key names
        if len(key_label) > 6:
            key_label = key_label[:5] + "."
        canvas_obj.drawString(x + 0.03 * inch, y - 0.12 * inch, key_label)
        
        # Decimal and Hex codes
        canvas_obj.setFont("Helvetica", 6)
        canvas_obj.setFillColorRGB(0.5, 0.5, 0.5)
        canvas_obj.drawString(x + 0.03 * inch, y - 0.20 * inch, f"{code}")
        canvas_obj.drawString(x + 0.03 * inch, y - 0.27 * inch, f"x{code:02X}")
        
        # Draw the character from the ESRI font (centered)
        try:
            canvas_obj.setFont(font_id, font_size_char)
            canvas_obj.setFillColorRGB(0, 0, 0)
            char = chr(code)
            canvas_obj.drawCentredString(x + cell_size/2, y - cell_size + 0.15 * inch, char)
            char_count += 1
        except:
            # If character can't be rendered
            canvas_obj.setFont("Helvetica", 8)
            canvas_obj.setFillColorRGB(0.8, 0.8, 0.8)
            canvas_obj.drawCentredString(x + cell_size/2, y - cell_size + 0.15 * inch, "—")
        
        canvas_obj.setFillColorRGB(0, 0, 0)
    
    return char_count

def draw_extended_grid(canvas_obj, font_id, font_name, start_code=256, end_code=512):
    """Draw extended Unicode characters if they exist"""
    y_start = page_height - 0.9 * inch
    x_start = margin
    
    char_count = 0
    for code in range(start_code, min(end_code + 1, start_code + (rows * cols))):
        row = (code - start_code) // cols
        col = (code - start_code) % cols
        
        if row >= rows:
            break
        
        x = x_start + (col * cell_size)
        y = y_start - (row * cell_size)
        
        canvas_obj.setStrokeColorRGB(0.8, 0.8, 0.8)
        canvas_obj.rect(x, y - cell_size, cell_size, cell_size)
        
        # Unicode label
        canvas_obj.setFont("Helvetica-Bold", font_size_label)
        canvas_obj.setFillColorRGB(0.2, 0.2, 0.2)
        canvas_obj.drawString(x + 0.03 * inch, y - 0.12 * inch, f"U+{code:04X}")
        
        # Decimal code
        canvas_obj.setFont("Helvetica", 6)
        canvas_obj.setFillColorRGB(0.5, 0.5, 0.5)
        canvas_obj.drawString(x + 0.03 * inch, y - 0.20 * inch, f"{code}")
        
        try:
            canvas_obj.setFont(font_id, font_size_char)
            canvas_obj.setFillColorRGB(0, 0, 0)
            char = chr(code)
            canvas_obj.drawCentredString(x + cell_size/2, y - cell_size + 0.15 * inch, char)
            char_count += 1
        except:
            canvas_obj.setFont("Helvetica", 8)
            canvas_obj.setFillColorRGB(0.8, 0.8, 0.8)
            canvas_obj.drawCentredString(x + cell_size/2, y - cell_size + 0.15 * inch, "—")
        
        canvas_obj.setFillColorRGB(0, 0, 0)
    
    return char_count

# Process each font
page_count = 0
for font_idx, font_name in enumerate(fonts):
    font_path = font_dict[font_name]
    
    try:
        # Register the font
        font_id = font_name.replace(" ", "_").replace("-", "_").replace(".", "_")
        
        if font_id not in pdfmetrics.getRegisteredFontNames():
            pdfmetrics.registerFont(TTFont(font_id, font_path))
        
        print(f"Processing {font_idx + 1}/{total_fonts}: {font_name}")
        
        # Page 1: ASCII range (32-255)
        page_count += 1
        if page_count > 1:
            c.showPage()
        
        draw_page_header(c, font_name, page_count, total_fonts * 2)
        
        # Add legend
        c.setFont("Helvetica", 9)
        c.drawString(margin, page_height - 0.75 * inch, 
                    "Each cell shows: Key name (top), Decimal code, Hex code, Symbol (center)")
        
        draw_character_grid(c, font_id, font_name, 32, 255)
        
        # Page 2: Extended Unicode range (256-511)
        page_count += 1
        c.showPage()
        
        draw_page_header(c, font_name, page_count, total_fonts * 2)
        
        c.setFont("Helvetica", 9)
        c.drawString(margin, page_height - 0.75 * inch,
                    "Extended Unicode (256-511) - Each cell shows: Unicode ID (top), Decimal code, Symbol (center)")
        
        draw_extended_grid(c, font_id, font_name, 256, 511)
        
    except Exception as e:
        print(f"  ERROR processing {font_name}: {str(e)}")
        continue

# Save PDF
c.save()

print("\n" + "=" * 70)
print("SUCCESS!")
print("=" * 70)
print(f"Character catalog created: {output_pdf}")
print(f"Total fonts processed: {total_fonts}")
print(f"Total pages: {page_count}")
print("=" * 70)
print("\nEach font has 2 pages:")
print("  Page 1: ASCII characters (codes 32-255)")
print("  Page 2: Extended Unicode (codes 256-511)")
print("  Format: Character shown with decimal and hex codes")
print("=" * 70)
