from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from cv2 import cv2
# Normal
# BodyText
# Italic
# Heading1
# Title
# Heading2
# Heading3
# Heading4
# Heading5
# Heading6
# Bullet
# Definition
# Code
# UnorderedList
# OrderedList

def link(link_text, show_text):
    link = '<a href="%(link)s" color="blue">%(show)s</a>'%{'link': link_text, 'show': show_text}
    return link

def generate_PDF(liveinfos, img_timetable, img_producttable):
    doc = SimpleDocTemplate("livetoday.pdf")

    im_timetable = cv2.imread(img_timetable)
    im_producttable = cv2.imread(img_producttable)
    h_time, w_time, c = im_timetable.shape
    h_prod, w_prod, c = im_producttable.shape
    c = None
    h_time = 150
    w_time = 500
    h_prod = 400
    w_prod = 500
    timetable_image = Image(img_timetable, w_time, h_time)
    producttable_image = Image(img_producttable, w_prod, h_prod)

    pdfmetrics.registerFont(TTFont('SimSun', './font/SimSun.ttf'))
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(fontName='SimSun', name='Song', leading=20, fontSize=12))
    story = []
    story.append(timetable_image)
    story.append(producttable_image)
    for live in liveinfos:
        story.append(Paragraph(live.influencer, styles['Heading2']))
        story.append(Paragraph(live.account, styles['Normal']))
        i = 0
        links = ''

        story.append(Paragraph('Infos :', styles['Heading5']))
        for code in live.codes:
            story.append(Paragraph(str(live.codes[code]), styles['Song']))

        story.append(Paragraph('Products :', styles['Heading5']))
        for product in live.products:
            story.append(Paragraph(link(live.products[product], product), \
                styles['Normal']))
            if i < len(live.products) -1 :
                links = links + live.products[product] + ","
            else :
                links = links + live.products[product]
            i += 1

        story.append(Paragraph('Replacement :', styles['Heading5']))
        story.append(Paragraph(links, styles['Italic']))

    doc.build(story)
