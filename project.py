import sys
from datetime import date
import docx


def main():
    if len(sys.argv) == 1:
        create_document()
    else:
        sys.exit(
            """\
        project.py - Generate a Word document that serves as a template for creating formal letters.
        
        Usage: py.exe project.py letter <company> - Generate template with company's address.
        py.exe project.py letter Committee - Generate template with standard committee address.
        """
        )


def create_document():
    ...


def date_doc():
    return date.today().strftime("%B %#d, %Y")


def address():
    ...


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
