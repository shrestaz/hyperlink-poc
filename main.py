import docx
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.opc.oxml import qn
from docx.oxml.shared import OxmlElement
from docx.oxml.shared import qn as qn_shared


def add_hyperlink(paragraph, url, text):
    """
    A function that places a hyperlink within a paragraph object.

    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :return: The hyperlink object
    """

    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id, )
    hyperlink.set(qn_shared("w:tgtFrame"), "_blank")

    # Create a w:r element
    new_run = OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = OxmlElement('w:rPr')

    r_style_element = OxmlElement("w:rStyle")
    r_style_element.set(qn_shared("w:val"), "InternetLink")
    rPr.append(r_style_element)

    text_element = OxmlElement("w:t")
    text_element.text = text

    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.append(text_element)
    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink



document = docx.Document()
p = document.add_paragraph()
add_hyperlink(p, 'https://www.google.com', 'Google')
document.save('demo.docx')