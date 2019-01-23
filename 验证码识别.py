#! /usr/bin/env python3
# -*- coding:utf-8 -*-

import tesserocr
from PIL import Image

image = Image.open('捷顺停车/code.jpg')
image = image.convert('L')
threshold = 60
table = []
for i in range(256):
	if i < threshold:
		table.append(0)
	else:
		table.append(1)
image = image.point(table, '1')
# image.show()
result = tesserocr.image_to_text(image)
print(result)
print(tesserocr.tesseract_version())
print(tesserocr.get_languages())