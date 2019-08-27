import openpyxl
import os
import subprocess
from shutil import copytree, make_archive
from tempfile import TemporaryDirectory
from zipfile import ZipFile


# Prep the template directory
try:
    template_dir = TemporaryDirectory()
    with ZipFile("template.docx") as template_doc:
        extracted_dir = template_doc.extractall(path=template_dir.name)
        print("Extracted to {}".format(extracted_dir))

    # Load the attendees list
    wb = openpyxl.load_workbook("people.xlsx")
    sheet = wb.get_sheet_by_name("Attendees")

    for row in sheet.iter_rows(
        min_row=2, max_row=100, max_col=6, values_only=True
    ):
        email = row[1]
        number = row[3] if row[3] is not None else ""
        name = row[4]
        points = row[5]
        if number.startswith("OT"):
            designation = "Occupational Therapist"
        elif number.startswith("PT"):
            designation = "Physiotherapist"
        elif number.startswith("ST"):
            designation = "Speech Therapist"
        else:
            designation = ""
        if not name:
            break

        print("{}  {}".format(name, email))

        # Create the rendered directory
        with TemporaryDirectory() as d:
            doc_dir = os.path.join(d, "doc")
            copytree(template_dir.name, doc_dir)

            # Render the variables
            infile = os.path.join(doc_dir, "word", "document.xml")
            outfile = os.path.join(doc_dir, "word", "new.xml")
            with open(infile, "rt") as f:
                with open(outfile, "wt") as g:
                    for line in f:
                        line = line.replace("##number##", number)
                        line = line.replace("##name##", name)
                        line = line.replace("##points##", str(points))
                        line = line.replace("##designation##", designation)
                        g.write(line)
            os.rename(outfile, infile)

            # Reassemble the rendered file
            # TODO name based filename
            newdoc = os.path.join(d, ".".join([name.replace(" ", ""), "docx"]))
            make_archive(newdoc, "zip", doc_dir)
            os.rename('.'.join([newdoc, "zip"]), newdoc)

            # Render the PDF
            with open(os.devnull, 'w') as null:
                subprocess.run(
                    [
                        "soffice",
                        "--headless",
                        "--invisible",
                        "--nocrashreport",
                        "--nodefault",
                        "--nologo",
                        "--nofirststartwizard",
                        "--norestore",
                        "--convert-to",
                        "pdf",
                        "--outdir",
                        ".",
                        newdoc,
                    ],
                    stdout=null,
                    stderr=null,
                )

finally:
    template_dir.cleanup()
