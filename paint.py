from PIL import Image, ImageDraw, ImageFont
import xlrd
import os

path = "C:/Users/KRITHIN/Desktop/Certificate-Generator-master/response.xls"
inputWorkbook = xlrd.open_workbook(path)
inputWorksheet = inputWorkbook.sheet_by_index(0)
rows = inputWorksheet.nrows

user = []
objects = {}
for i in range(rows):
    objects['name'] = inputWorksheet.cell_value(i, 1)
    objects['usn'] = inputWorksheet.cell_value(i, 2)
    user.append(objects)
    objects = {}


for person in user:
    image = Image.open('Capture.jpg')
    name = person['name']
    usn = person['usn']
    font = ImageFont.truetype("arial.ttf", 35)

    draw = ImageDraw.Draw(image)
    draw.text(xy=(280, 505), text=usn, fill=(0,0,0), font=font)
    draw.text(xy=(510, 505), text=name, fill=(0, 0, 0), font=font)
    image.save(f'{os.getcwd()}/certificates/{usn}_{name}.jpg')
    print(f'Generated for {usn}_{name}')
