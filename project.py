import os
import sys
from datetime import date
import docx


def main():
    os.chdir(r"C:\Users\HP\PycharmProjects\CS50-Final-Project")
    if len(sys.argv) == 4:
        create_document()
    else:
        sys.exit(
            """\
        project.py - Generate a Word document that serves as a template for creating formal letters.
        
        Usage: py.exe project.py letter to <company> - Generate template with company's address.
        py.exe project.py letter to committee - Generate template with standard committee address.
        """
        )


def create_document():
    doc_name = f"{sys.argv[1]} {sys.argv[2]} {sys.argv[3]}.docx"
    doc = docx.Document()
    doc.add_paragraph(draft())
    doc.save(doc_name)


def date_doc():
    return date.today().strftime("%B %#d, %Y")


def address():
    return "The Managing Director\nVolta River Authority\nElectro-Volta House\nAccra"


def reference():
    ref_year = str(date.today().year)[2:]
    if sys.argv[2] == "committee":
        return "REF:EC/LCLP/COM/" + ref_year + "/0.."
    else:
        ref_year = str(date.today().year)[2:]
        return "EC/LCLP/EMO/" + ref_year + "/0.."


def draft():
    return "DRAFT"


def greeting():
    return "Dear Sir/Madam,"


def title():
    return "TITLE:"


def complimentary_close():
    return "Yours faithfully,\n\nIng.Oscar Amonoo-Neizer\n(Executive Secretary)"


if __name__ == "__main__":
    main()
