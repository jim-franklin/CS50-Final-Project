import os
import shelve
import sys
from datetime import date

import pyperclip
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_LINE_SPACING
from docx.shared import Pt
from docx.shared import Inches


def main():
    os.chdir(r"C:\Users\HP\PycharmProjects\CS50-Final-Project")
    letters = ["letter", "let"]
    lists = ["list", "ls"]

    if len(sys.argv) == 4:
        if sys.argv[1].lower() not in letters:
            raise TypeError(f"\n\n**Did you mean to letter to {sys.argv[3]}?**\n\n")
        create_document()

    elif len(sys.argv) == 2 and sys.argv[1].lower() == "save":
        save_address()

    elif len(sys.argv) == 2 and sys.argv[1].lower() in lists:
        list_all_address()

    else:
        sys.exit(
            """\
        project.py - Generate a Word document that serves as a template for creating formal letters.
        
        Usage: py.exe project.py letter to <company> - Generate template with company's address.
        py.exe project.py letter to committee - Generate template with standard committee address.
        """
        )


def create_document():
    doc = Document()

    # Set custom style, font, font size and alignment for entire document
    base_style = doc.styles["Normal"]
    base_style.font.name = "Times New Roman"
    base_style.font.size = Pt(12)
    base_style.paragraph_format.alignment = WD_PARAGRAPH_ALIGNMENT.JUSTIFY

    # Set custom margins
    section = doc.sections[0]
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    # Add Draft
    p_draft = doc.add_paragraph(draft())
    p_draft.runs[0].bold = True
    p_draft.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Add Reference
    p_reference = doc.add_paragraph(reference())
    p_reference.runs[0].bold = True
    p_reference.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add Date
    p_date_doc = doc.add_paragraph(date_doc())
    p_date_doc.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT

    # Add Address
    p_address = doc.add_paragraph(address())
    p_address.runs[0].bold = True
    p_address.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add Greeting or not if the letter is to the committee
    if sys.argv[3].lower() != "committee":
        para_greet = doc.add_paragraph(greeting())
        para_greet.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add Title
    p_title = doc.add_paragraph(title())
    p_title.runs[0].underline = True
    p_title.runs[0].bold = True
    p_title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    # Add body of letter
    body_body = body()
    # body of committee letter
    if sys.argv[3].lower() == "committee":
        string = body_body.split("\n")
        for line in string:
            p_body = doc.add_paragraph(line)
            p_body.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
    # body of any other letter
    else:
        p_body = doc.add_paragraph(body_body)
        p_body.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE

    # Add valediction
    p_valediction = doc.add_paragraph(valediction())
    p_valediction.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Add name of writer
    p_writer_name = doc.add_paragraph(writer_name())
    p_writer_name.runs[0].bold = True
    p_writer_name.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

    # Save Document
    document_name = f"{sys.argv[1]}_{sys.argv[2]}_{sys.argv[3]}.docx"
    doc.save(document_name)
    print(document_name, "has been created...\n")


def date_doc():
    return date.today().strftime("%B %#d, %Y")


def address():
    if sys.argv[3].lower() == "committee":
        return "CONFIDENTIAL\n\nTO THE COMMITTEE MEMBERS\nOF THE LOCAL CONTENT AND LOCAL PARTICIPATION\nCOMMITTEE"
    else:
        company_address()


def reference():
    ref_year = str(date.today().year)[2:]
    if sys.argv[3].lower() == "committee":
        return "REF: EC/LCLP/COM/" + ref_year + "/0.."
    else:
        ref_year = str(date.today().year)[2:]
        return "EC/LCLP/EMO/" + ref_year + "/0.."


def draft():
    return "DRAFT"


def greeting():
    return "Dear Sir/Madam,"


def title():
    if sys.argv[3].lower() == "committee":
        return (
            "INVITATION TO MEETING FOR LOCAL CONTENT AND\nLOCAL PARTICIPATION COMMITTEE"
        )
    return "ENTER TITLE HERE:"


def body():
    if sys.argv[3].lower() == "committee":
        date_of_meeting = input(
            "Date of the meeting (e.g. Tuesday, March  7, 2023 and Wednesday, March 8 2023): "
        )
        time_of_meeting = input("Time of the meeting (e.g. 10:00 am): ")
        return (
            f"Members are kindly invited to a virtual meeting of the Local Content and Local Participation "
            f"Committee on {date_of_meeting}, at {time_of_meeting}.\nThe agenda for the meeting is as follows:\n\t1. "
        )
    return "..."


def valediction():
    top_space = "\n" * 10
    return top_space + "Yours faithfully,\n"


def writer_name():
    return "Ing. Oscar Amonoo-Neizer\n(Executive Secretary)"


def save_address():
    with shelve.open("data_base") as sfile:
        name_of_company = input(
            "Enter the name of the company whose address you want to save: "
        )
        sfile[name_of_company] = input(
            "Enter the company's address:\n(e.g. The Managing Director, Volta River Authority, Electro-Volta House, "
            "Accra)"
        )
        if not sfile[name_of_company]:
            sfile[name_of_company] = pyperclip.paste()


def list_all_address():
    with shelve.open("data_base") as sfile:
        print(list(sfile.keys()))


def company_address():
    with shelve.open("data_base") as sfile:
        if sys.argv[3] in sfile:
            return sfile[sys.argv[3]]
        else:
            save_address()


if __name__ == "__main__":
    main()
