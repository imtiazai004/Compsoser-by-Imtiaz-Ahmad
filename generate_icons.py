from PIL import Image, ImageDraw
import os

def create_icon(size, path):
    img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    r = int(size * 0.22)
    draw.rounded_rectangle([0, 0, size-1, size-1], radius=r, fill=(79, 70, 229, 255))

    pad  = int(size * 0.18)
    fold = int(size * 0.20)
    dx1, dy1 = pad, pad
    dx2, dy2 = size - pad, size - pad

    draw.polygon([
        (dx1, dy1 + fold),
        (dx2 - fold, dy1),
        (dx2, dy1 + fold),
        (dx2, dy2),
        (dx1, dy2),
    ], fill='white')

    draw.polygon([
        (dx2 - fold, dy1),
        (dx2, dy1 + fold),
        (dx2 - fold, dy1 + fold),
    ], fill=(196, 181, 253))

    h  = dy2 - dy1
    lx1 = dx1 + int(size * 0.08)
    lx2 = dx2 - int(size * 0.08)
    lh  = max(2, int(size * 0.04))
    lc  = (167, 139, 250, 200)

    for i, frac in enumerate([0.50, 0.63, 0.76]):
        ly = dy1 + int(h * frac)
        ex = lx2 if i < 2 else lx1 + int((lx2 - lx1) * 0.60)
        draw.rectangle([lx1, ly, ex, ly + lh], fill=lc)

    img.save(path)
    print(f"icon {size}x{size} -> {path}")

if __name__ == '__main__':
    os.makedirs('static', exist_ok=True)
    create_icon(192, 'static/icon-192.png')
    create_icon(512, 'static/icon-512.png')
