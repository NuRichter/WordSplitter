"""Draw assets/wordsplitter.ico (developer tool, needs Pillow).

The mark is a Word-blue tile holding a page that is cut in two, the upper
half in white and the lower half shifted aside, with a dashed cut line.
Run:  py -3 tools\\make_icon.py
"""

from pathlib import Path

from PIL import Image, ImageDraw

ROOT = Path(__file__).resolve().parent.parent
SIZE = 1024
BLUE = (24, 90, 189, 255)
TEAL = (14, 124, 102, 255)
WHITE = (255, 255, 255, 255)


def draw() -> Image.Image:
    image = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
    d = ImageDraw.Draw(image)
    d.rounded_rectangle((40, 40, SIZE - 40, SIZE - 40), radius=210, fill=BLUE)
    # Upper half of the page.
    d.rounded_rectangle((270, 190, 700, 500), radius=40, fill=WHITE)
    # Lower half, shifted to show the split.
    d.rounded_rectangle((330, 560, 760, 840), radius=40, fill=(220, 240, 236, 255))
    for y, color in ((270, BLUE), (350, BLUE), (430, BLUE)):
        d.rounded_rectangle((330, y - 18, 640, y + 18), radius=18, fill=color)
    for y in (640, 720):
        d.rounded_rectangle((390, y - 18, 700, y + 18), radius=18, fill=TEAL)
    # Dashed cut line across the tile.
    x = 150
    while x < SIZE - 150:
        d.rounded_rectangle((x, 515, x + 70, 545), radius=15, fill=WHITE)
        x += 110
    return image


def main() -> None:
    target = ROOT / "assets" / "wordsplitter.ico"
    target.parent.mkdir(parents=True, exist_ok=True)
    image = draw()
    image.save(target, sizes=[(16, 16), (20, 20), (24, 24), (32, 32), (40, 40), (48, 48), (64, 64), (128, 128), (256, 256)])
    image.resize((256, 256), Image.LANCZOS).save(ROOT / "assets" / "wordsplitter.png")
    print("written", target)


if __name__ == "__main__":
    main()
