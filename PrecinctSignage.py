from pptx import Presentation
from pptx.util import Inches
import csv
import datetime
import time

# define inputs
file_location = "C:\\Users\\amalc\\OneDrive\\Resistor\\Code\\PrecinctSignage\\" #double backslashes to escape the escapes
preso_name = "PrecinctSignageTemplate.pptx"
csvinput_name = "PrecinctSignageTemplate.csv"

# create the presentation
prs = Presentation(file_location + preso_name)

title_slide_layout = prs.slide_layouts[1]
signage_slide_layout = prs.slide_layouts[2]

slide1 = prs.slides.add_slide(title_slide_layout)
title = slide1.shapes.title
subtitle = slide1.placeholders[1]

title.text = "Precinct Signage" 
subtitle.text= "run at " + datetime.datetime.now().strftime("%I:%M%p on %B %d, %Y")


with open(file_location + csvinput_name) as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        #add a slide
        slide = prs.slides.add_slide(signage_slide_layout)

        # define the placeholders
        organization_name = slide.shapes.title
        election_name = slide.placeholders[1]
        precinct_name = slide.placeholders[2]
        race1_name = slide.placeholders[3]
        race1_gfx = slide.placeholders[4]
        race2_name = slide.placeholders[5]
        race2_gfx = slide.placeholders[6]
        race3_name = slide.placeholders[7]
        race3_gfx = slide.placeholders[8]
        race4_name = slide.placeholders[9]
        race4_gfx = slide.placeholders[10]
        race5_name = slide.placeholders[11]
        race5gfx = slide.placeholders[12]
        race6_name = slide.placeholders[13]
        race6_gfx = slide.placeholders[14]

        # fill the placeholders
        organization_name.txt = row['organization']
        election_name.txt = row['election_name']
        precinct_name.txt = row['precinct_name']
        race1_name.txt = row['race1_name']
        race1pic = race1_gfx.insert_picture(file_location + row['race1_gfx'])
        race2_name.txt = row['race2_name']
        race2pic = race2_gfx.insert_picture(file_location + row['race2_gfx'])
        race3_name.txt = row['race3_name']
        race3pic = race3_gfx.insert_picture(file_location + row['race3_gfx'])
        race4_name.txt = row['race4_name']
        race4pic = race4_gfx.insert_picture(file_location + row['race4_gfx'])
        race5_name.txt = row['race5_name']
        race5pic = race5_gfx.insert_picture(file_location + row['race5_gfx'])
        race6_name.txt = row['race6_name']
        race6pic = race6_gfx.insert_picture(file_location + row['race6_gfx'])

# When done save file
timestr = time.strftime("%Y%m%d-%H%M%S")
filenamestr = "Precinct Signage" + timestr
prs.save(filenamestr)