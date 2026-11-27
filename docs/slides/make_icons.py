"""Renders the slide icons from the Lucide font the patient app bundles, so the
slides use the app's own icons. Writes icons/<name>-<colour>.png (256 px).

    python make_icons.py

Needs Pillow and the lucide_icons_flutter package in the Flutter pub cache.
"""
import glob
import json
import os

from PIL import Image, ImageDraw, ImageFont

NAMES = [
    'repeat', 'pill', 'users', 'lock', 'timer', 'cloud-off', 'sparkles', 'scan-text', 'stethoscope',
    'clipboard-check', 'accessibility', 'smartphone', 'shield-check', 'server', 'key-round', 'database',
    'git-branch', 'test-tube', 'circle-check', 'triangle-alert', 'globe', 'languages', 'hospital',
    'file-text', 'workflow', 'layers', 'badge-check', 'heart-pulse', 'qr-code', 'code', 'scan-line',
    'hand-heart', 'eye-off', 'book-open', 'list-checks', 'gauge', 'user-round', 'activity', 'target',
]
COLOURS = {'navy': (0, 32, 96, 255), 'white': (255, 255, 255, 255), 'gold': (255, 202, 8, 255)}

cache = os.path.join(os.environ['LOCALAPPDATA'], 'Pub', 'Cache', 'hosted', 'pub.dev')
pkg = sorted(glob.glob(os.path.join(cache, 'lucide_icons_flutter-*')))[-1]
codes = json.load(open(os.path.join(pkg, 'assets', 'codepoints.json'), encoding='utf-8'))
font = ImageFont.truetype(os.path.join(pkg, 'assets', 'lucide.ttf'), 200)
out = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'icons')
os.makedirs(out, exist_ok=True)
for name in NAMES:
    glyph = chr(codes[name])
    for cname, rgba in COLOURS.items():
        im = Image.new('RGBA', (256, 256), (0, 0, 0, 0))
        d = ImageDraw.Draw(im)
        l, t, r, b = d.textbbox((0, 0), glyph, font=font)
        d.text(((256 - (r - l)) / 2 - l, (256 - (b - t)) / 2 - t), glyph, font=font, fill=rgba)
        im.save(os.path.join(out, f'{name}-{cname}.png'))
print(len(NAMES), 'icons from', os.path.basename(pkg))
