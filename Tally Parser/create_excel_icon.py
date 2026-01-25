#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Script to create a crossed-out Excel-like icon
"""

from PIL import Image, ImageDraw, ImageFont
import os

def create_crossed_excel_icon(size=256, output_path="excel_crossed.png"):
    """
    Create an Excel-like icon with a diagonal cross line
    
    Args:
        size: Size of the icon in pixels (default 256)
        output_path: Path to save the icon
    """
    # Create image with transparent background
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    
    # Excel green color (#107C41)
    excel_green = (16, 124, 65)
    
    # Calculate dimensions
    padding = int(size * 0.1)  # 10% padding
    corner_radius = int(size * 0.15)  # Rounded corners
    
    # Draw rounded rectangle background (Excel-like)
    # We'll draw it as a rectangle with rounded corners manually
    # Top-left corner
    draw.ellipse([padding, padding, padding + corner_radius * 2, padding + corner_radius * 2], 
                 fill=excel_green)
    # Top-right corner
    draw.ellipse([size - padding - corner_radius * 2, padding, size - padding, padding + corner_radius * 2], 
                 fill=excel_green)
    # Bottom-left corner
    draw.ellipse([padding, size - padding - corner_radius * 2, padding + corner_radius * 2, size - padding], 
                 fill=excel_green)
    # Bottom-right corner
    draw.ellipse([size - padding - corner_radius * 2, size - padding - corner_radius * 2, size - padding, size - padding], 
                 fill=excel_green)
    
    # Fill the main rectangle
    draw.rectangle([padding, padding + corner_radius, size - padding, size - padding - corner_radius], 
                   fill=excel_green)
    draw.rectangle([padding + corner_radius, padding, size - padding - corner_radius, size - padding], 
                   fill=excel_green)
    
    # Draw white "X" in the center (Excel logo style)
    x_thickness = int(size * 0.08)
    x_size = int(size * 0.4)
    center_x, center_y = size // 2, size // 2
    
    # Draw the X (two diagonal lines)
    # Top-left to bottom-right
    draw.line([center_x - x_size//2, center_y - x_size//2, 
               center_x + x_size//2, center_y + x_size//2], 
              fill=(255, 255, 255), width=x_thickness)
    # Top-right to bottom-left
    draw.line([center_x + x_size//2, center_y - x_size//2, 
               center_x - x_size//2, center_y + x_size//2], 
              fill=(255, 255, 255), width=x_thickness)
    
    # Draw diagonal cross line (red, thicker) to indicate "crossed out"
    cross_thickness = int(size * 0.12)
    cross_padding = int(size * 0.15)
    
    # Red color for the cross
    red_cross = (220, 50, 47)
    
    # Draw diagonal line from top-left to bottom-right
    draw.line([cross_padding, cross_padding, 
               size - cross_padding, size - cross_padding], 
              fill=red_cross, width=cross_thickness)
    
    # Save the image
    img.save(output_path, 'PNG')
    print(f"Icon created: {output_path}")
    
    # Also create ICO format for Windows
    ico_path = output_path.replace('.png', '.ico')
    # Create multiple sizes for ICO
    sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
    ico_images = []
    for ico_size in sizes:
        resized = img.resize(ico_size, Image.Resampling.LANCZOS)
        ico_images.append(resized)
    
    ico_images[0].save(ico_path, format='ICO', sizes=[(s[0], s[1]) for s in sizes])
    print(f"ICO icon created: {ico_path}")

if __name__ == "__main__":
    # Get the directory where this script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Create icons in the script directory (project root)
    png_path = os.path.join(script_dir, "excel_crossed.png")
    create_crossed_excel_icon(size=256, output_path=png_path)
    
    print("\nIcons created successfully!")
    print(f"PNG: {png_path}")
    print(f"ICO: {png_path.replace('.png', '.ico')}")
