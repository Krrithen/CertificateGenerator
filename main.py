from PIL import Image, ImageDraw, ImageFont
import pandas as pd
form = pd.read_excel("participation.xls")
name_list = form['name'].to_list()
for i in name_list:
    im = Image.open("participation.jpg")
    d = ImageDraw.Draw(im)
    location = (760, 660)
    text_color = (0, 0, 0)
    font = ImageFont.truetype("arial.ttf", 100)
    d.text(location, i, fill=text_color,font=font)
    im.save("certificate_"+i+".pdf")