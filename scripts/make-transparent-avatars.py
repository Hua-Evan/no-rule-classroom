from collections import deque
from pathlib import Path

import numpy as np
from PIL import Image, ImageFilter


ROOT = Path(__file__).resolve().parents[1]
SOURCE = ROOT / "assets" / "avatars-grid.png"
TARGET = ROOT / "assets" / "avatars-grid-transparent.png"


def transparent_cell(cell: Image.Image) -> Image.Image:
    rgb = np.asarray(cell.convert("RGB"), dtype=np.int16)
    height, width, _ = rgb.shape
    border = np.concatenate((rgb[0], rgb[-1], rgb[:, 0], rgb[:, -1]))
    paper = np.median(border, axis=0)
    delta = np.sqrt(np.sum((rgb - paper) ** 2, axis=2))
    channel_range = rgb.max(axis=2) - rgb.min(axis=2)
    paper_like = (delta < 46) & (rgb.mean(axis=2) > 188) & (channel_range < 58)

    background = np.zeros((height, width), dtype=bool)
    queue = deque()
    for x in range(width):
        queue.extend(((0, x), (height - 1, x)))
    for y in range(height):
        queue.extend(((y, 0), (y, width - 1)))

    while queue:
        y, x = queue.popleft()
        if background[y, x] or not paper_like[y, x]:
            continue
        background[y, x] = True
        if y:
            queue.append((y - 1, x))
        if y + 1 < height:
            queue.append((y + 1, x))
        if x:
            queue.append((y, x - 1))
        if x + 1 < width:
            queue.append((y, x + 1))

    alpha = Image.fromarray((~background * 255).astype(np.uint8), "L")
    alpha = alpha.filter(ImageFilter.GaussianBlur(0.7))
    result = cell.convert("RGBA")
    result.putalpha(alpha)
    return result


sheet = Image.open(SOURCE).convert("RGB")
cell_width, cell_height = sheet.width // 4, sheet.height // 2
output = Image.new("RGBA", sheet.size, (0, 0, 0, 0))
for row in range(2):
    for column in range(4):
        box = (
            column * cell_width,
            row * cell_height,
            (column + 1) * cell_width,
            (row + 1) * cell_height,
        )
        output.paste(transparent_cell(sheet.crop(box)), box)

output.save(TARGET, optimize=True)
print(TARGET)
