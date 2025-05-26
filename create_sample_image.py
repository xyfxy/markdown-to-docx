from PIL import Image, ImageDraw

def create_image(filename="sample_image.png", size=(100, 50), color="blue"):
    """Creates a simple PNG image with a solid color."""
    try:
        img = Image.new('RGB', size, color=color)
        draw = ImageDraw.Draw(img)
        # Optional: Add some text to make it slightly more identifiable
        draw.text((10, 10), "Sample", fill="white")
        img.save(filename)
        print(f"'{filename}' created successfully ({size[0]}x{size[1]}, color: {color}).")
    except Exception as e:
        print(f"Error creating image '{filename}': {e}")

if __name__ == "__main__":
    create_image()
