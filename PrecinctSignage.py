from pptx import Presentation
from pptx.util import Inches
import csv

prs = Presentation("C:\Users\amalc\OneDrive\Resistor\Code\PrecinctSignage\PrecinctSignageTemplate.pptx")
title_slide_layout = prs.slide_layouts[1]
blank_slide_layout = prs.slide_layouts[2]




with open('names.csv') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        print(row['first_name'], row['last_name'])
