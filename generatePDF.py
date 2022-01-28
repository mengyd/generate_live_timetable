from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PIL import Image as Img
from loadconfig import loadConfig
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

# Load configurations
config = loadConfig(choice='params')

def link(show_text, link_text):
    link = '<a href="%(link)s" color="blue">%(show)s</a>'\
        %{'link': link_text, 'show': show_text}
    return link

def generate_PDF(liveinfos, img_timetable, img_producttable, filename):
    doc = SimpleDocTemplate(filename[4:]+".pdf")
    
    # Calculate image height
    im_timetable = Img.open(img_timetable)
    im_timetable_width, im_timetable_height = im_timetable.size


    image_width = config["table_image_width"]

    timetable_zoomrate = image_width/im_timetable_width

    h_time = im_timetable_height * timetable_zoomrate

    # Load images with defined sizes
    timetable_image = Image(img_timetable, image_width, h_time)

    pdfmetrics.registerFont(TTFont('SimSun', './font/SimSun.ttf'))
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(fontName='SimSun', name='Song', leading=20, fontSize=12))
    styles.add(ParagraphStyle(fontName='SimSun', name='Song_sm', leading=20, fontSize=10))
    styles.add(ParagraphStyle(fontName='SimSun', name='Song_bg', leading=20, fontSize=14))
    story = []
    story.append(timetable_image)

    if config["load_productImg"]:
        im_producttable = Img.open(img_producttable)
        im_producttable_width, im_producttable_height = im_producttable.size
        producttable_zoomrate = image_width/im_producttable_width
        h_prod = im_producttable_height * producttable_zoomrate
        producttable_image = Image(img_producttable, image_width, h_prod)
        story.append(producttable_image)

    for live in liveinfos:
        story.append(Paragraph(live.influencer, styles['Heading2']))
        story.append(Paragraph(live.account, styles['Normal']))
        i = 0
        links = ''
        ids = ''
        codes = []

        story.append(Paragraph('Infos :', styles['Heading5']))
        for product in live.products:
            for code in product.codes:
                if code not in codes:
                    codes.append(code)
                    story.append(Paragraph(str(code), styles['Song']))

        story.append(Paragraph('Products :', styles['Heading5']))
        for product in live.products:
            # 生成产品链接，参数产品名，产品链接
            story.append(Paragraph(link(product.name, product.link), \
                styles['Song']))
            story.append(Paragraph("产品名: " + product.alias \
                + ";\tID: " + product.id, styles['Song_sm']))
            if product.codes:
                story.append(Paragraph("Infos: ", styles['Bullet']))
                for code in product.codes:
                    story.append(Paragraph(code, styles['Song_sm']))
            # 产品ID
            if i < len(live.products) -1 :
                ids = ids + product.id + ","
            else :
                ids = ids + product.id
            i += 1
            story.append(Paragraph('-'*25, styles['Heading2']))

        story.append(Paragraph('Products ID :', styles['Heading5']))
        story.append(Paragraph(ids, styles['Italic']))
        story.append(Paragraph('='*40, styles['Heading1']))

    doc.build(story)
