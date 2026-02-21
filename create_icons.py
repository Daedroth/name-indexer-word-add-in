#!/usr/bin/env python3
from PIL import Image, ImageDraw

# Create Settings icon (gear) - 32x32
img_settings = Image.new('RGBA', (32, 32), (255, 255, 255, 0))
draw = ImageDraw.Draw(img_settings)

# Draw gear - settings icon
# Center circle
draw.ellipse([12, 12, 20, 20], fill='#0078d4')

# Gear teeth
draw.rectangle([14, 2, 18, 8], fill='#0078d4')      # Top
draw.rectangle([14, 24, 18, 30], fill='#0078d4')    # Bottom
draw.rectangle([2, 14, 8, 18], fill='#0078d4')      # Left
draw.rectangle([24, 14, 30, 18], fill='#0078d4')    # Right
draw.rectangle([21, 6, 25, 10], fill='#0078d4')     # Top-right
draw.rectangle([21, 22, 25, 26], fill='#0078d4')    # Bottom-right
draw.rectangle([7, 6, 11, 10], fill='#0078d4')      # Top-left
draw.rectangle([7, 22, 11, 26], fill='#0078d4')     # Bottom-left

img_settings.save('assets/icon-settings.png')
print("✓ Settings icon created (32x32)")

# For larger sizes
for size in [16, 80]:
    img = img_settings.resize((size, size), Image.Resampling.LANCZOS)
    img.save(f'assets/icon-settings-{size}.png')
    print(f"✓ Settings icon created ({size}x{size})")

# Create Index icon (document with index marker) - 32x32
img_index = Image.new('RGBA', (32, 32), (255, 255, 255, 0))
draw = ImageDraw.Draw(img_index)

# Document rectangle
draw.rectangle([5, 3, 21, 27], outline='#0078d4', width=2)

# Text lines on document
draw.line([(8, 9), (18, 9)], fill='#0078d4', width=1)
draw.line([(8, 14), (18, 14)], fill='#0078d4', width=1)
draw.line([(8, 19), (14, 19)], fill='#0078d4', width=1)

# Index flag/marker
flag_points = [(22, 8), (28, 4), (28, 12)]
draw.polygon(flag_points, fill='#0078d4')

img_index.save('assets/icon-index.png')
print("✓ Index icon created (32x32)")

# For larger sizes
for size in [16, 80]:
    img = img_index.resize((size, size), Image.Resampling.LANCZOS)
    img.save(f'assets/icon-index-{size}.png')
    print(f"✓ Index icon created ({size}x{size})")

print("\nAll icons created successfully!")
