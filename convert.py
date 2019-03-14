from PIL import Image
import xlsxwriter

import os
import sys

from itertools import islice

def chunks(l, n):
    """Yield successive n-sized chunks from l."""
    for i in range(0, len(l), n):
        yield l[i:i + n]


def color_string(rgb):
    return "#" + "".join(map(lambda c: "%0.2X" % c, rgb))


if len(sys.argv) < 2:
    print("usage: python %s <image>" % sys.argv[0])
    exit(1)

image_path = sys.argv[1]
im = Image.open(image_path)

width, height = im.size
px = im.load()

n_colors = max(set(im.getdata())) + 1
colors   = list(islice(chunks(im.getpalette(), 3), n_colors))

out_path = os.path.join('outputs', 'demo.xlsx')

workbook  = xlsxwriter.Workbook(out_path)
worksheet = workbook.add_worksheet()

worksheet.set_column(0, width, 1.5)

for x in range(width):
    for y in range(height):
        p = px[x, y]
        color = color_string(colors[p])

        cell_format = workbook.add_format()
        cell_format.set_bg_color(color)

        worksheet.write(y, x, '', cell_format)

workbook.close()
