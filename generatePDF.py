from reportlab.platypus import SimpleDocTemplate, Paragraph, Image
from reportlab.lib.styles import getSampleStyleSheet
import reportlab, LiveInfo, cv2
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

def generate_html(liveinfos, img_timetable, img_producttable):
    doc = SimpleDocTemplate("livetoday.pdf")
    im_timetable = cv2.imread(img_timetable)
    im_producttable = cv2.imread(img_producttable)
    h_time, w_time, c = im_timetable.shape
    h_prod, w_prod, c = im_producttable.shape
    h_time = h_time*0.15
    w_time = w_time*0.15
    h_prod = h_prod*0.25
    w_prod = w_prod*0.25
    timetable_image = Image(img_timetable, w_time, h_time)
    producttable_image = Image(img_producttable, w_prod, h_prod)
    styles = getSampleStyleSheet()
    story = []
    story.append(timetable_image)
    story.append(producttable_image)
    for live in liveinfos:
        story.append(Paragraph(live.influencer, styles['Heading2']))
        story.append(Paragraph(live.account, styles['Normal']))
        i = 0
        links = ''
        for product in live.products:
            story.append(Paragraph(link(live.products[product], product), styles['Normal']))
            if i < len(live.products) -1 : 
                links = links + live.products[product] + ","
            else :
                links = links + live.products[product]
            i += 1
        story.append(Paragraph('Replacement :', styles['Heading5']))
        story.append(Paragraph(links, styles['Italic']))


        # story.append(Paragraph(live.codes, style))

    # html = """
    #     <html>
    #     <head></head>
    #     <body>
    #     <p>Hello,World!</p>
    #     <p>Add webbrowser function</p>
    #     <p>%s</p>
    #     <p>%s</p>
    #     </body>
    #     </html>"""%("str_1,str_2", "")

    # return html
    doc.build(story)

