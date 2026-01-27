"""Generate a modern iC icon for iCharlotte application."""

from PIL import Image, ImageDraw, ImageFont
import os

def create_icon():
    """Create a modern iC icon with multiple sizes for .ico file."""

    sizes = [16, 24, 32, 48, 64, 128, 256]
    images = []

    for size in sizes:
        # Create image with transparent background
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        # Modern gradient-style solid color - deep blue to teal gradient feel
        # Using a rich blue color (#1565C0 - matches app theme)
        primary_color = (21, 101, 192)  # Deep blue
        accent_color = (0, 200, 200)     # Teal accent

        # Draw rounded rectangle background
        padding = max(1, size // 16)
        corner_radius = max(2, size // 6)

        # Draw the rounded rectangle
        draw.rounded_rectangle(
            [padding, padding, size - padding - 1, size - padding - 1],
            radius=corner_radius,
            fill=primary_color
        )

        # Add subtle inner glow/highlight at top
        highlight_height = max(1, size // 8)
        for i in range(highlight_height):
            alpha = int(60 * (1 - i / highlight_height))
            highlight_color = (255, 255, 255, alpha)
            y = padding + corner_radius // 2 + i
            x_start = padding + corner_radius // 2
            x_end = size - padding - corner_radius // 2
            if y < size // 3:
                draw.line([(x_start, y), (x_end, y)], fill=highlight_color)

        # Calculate font size - make text prominent
        font_size = int(size * 0.55)

        # Try to use a nice font, fall back to default
        font = None
        font_paths = [
            "C:/Windows/Fonts/segoeuib.ttf",   # Segoe UI Bold
            "C:/Windows/Fonts/segoeui.ttf",    # Segoe UI
            "C:/Windows/Fonts/arialbd.ttf",    # Arial Bold
            "C:/Windows/Fonts/arial.ttf",      # Arial
        ]

        for font_path in font_paths:
            if os.path.exists(font_path):
                try:
                    font = ImageFont.truetype(font_path, font_size)
                    break
                except:
                    continue

        if font is None:
            font = ImageFont.load_default()

        # Draw "iC" text
        text = "iC"

        # Get text bounding box for centering
        bbox = draw.textbbox((0, 0), text, font=font)
        text_width = bbox[2] - bbox[0]
        text_height = bbox[3] - bbox[1]

        # Center the text
        x = (size - text_width) // 2 - bbox[0]
        y = (size - text_height) // 2 - bbox[1]

        # Draw text shadow for depth (only on larger sizes)
        if size >= 32:
            shadow_offset = max(1, size // 32)
            draw.text((x + shadow_offset, y + shadow_offset), text,
                     fill=(0, 0, 0, 80), font=font)

        # Draw main text in white
        draw.text((x, y), text, fill=(255, 255, 255), font=font)

        images.append(img)

    # Save as .ico file with all sizes
    icon_path = os.path.join(os.path.dirname(__file__), 'icharlotte.ico')

    # Save the largest image as the base, include all sizes
    images[-1].save(
        icon_path,
        format='ICO',
        sizes=[(s, s) for s in sizes]
    )

    # Also save individual PNG for reference
    png_path = os.path.join(os.path.dirname(__file__), 'icharlotte_icon.png')
    images[-1].save(png_path, format='PNG')

    print(f"Icon saved to: {icon_path}")
    print(f"PNG preview saved to: {png_path}")
    return icon_path

if __name__ == "__main__":
    create_icon()
