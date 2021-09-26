from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (Flowable, Paragraph, SimpleDocTemplate,
                                Spacer, Table, TableStyle)
from reportlab.lib import colors
import PyPDF4
import re
import os


class LineDrw(Flowable):
    def __init__(self, width, height=0):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def __repr__(self):
        return "Line(w=%s)" % self.width

    def draw(self):
        self.canv.line(0, self.height, self.width, self.height)


class Report:
    def __init__(self, datatable, rptdata, colnames,
                 ColWdth, filename, title=None):
        # ( table name, table data, table column names, table column widths,
        # name of PDF file)
        self.columnames = colnames
        self.colwdths = ColWdth
        self.filename = filename
        self.rptdata = rptdata

        self.width, self.height = letter

        self.rptdata.insert(0, tuple(self.columnames))

        if title is None:
            if datatable.find("_") != -1:
                self.ttl = (datatable.replace("_", " "))
            else:
                self.ttl = (' '.join(re.findall('([A-Z][a-z]*)', datatable)))
        else:
            self.ttl = title

    def create_pdf(self):
        body = []
        textAdjust = 6.5

        if not self.filename:
            exit()

        styles = getSampleStyleSheet()
        spacer = Spacer(0, 0.25*inch)
        ptext = '<font size=14>%s</font>' % self.ttl
        body.append(Paragraph(ptext, styles["Heading2"]))

        line = LineDrw(500)

        body.append(line)
        body.append(spacer)

        colwidths = []
        for i in self.colwdths:
            colwidths.append(i * textAdjust)

        m = 0
        new_data = []
        for lin in list(self.rptdata):
            lin = list(lin)
            n = 0
            for item in lin:
                if type(item) == str:
                    if len(item) >= 10:
                        item = Paragraph(item, styles['Normal'])
                        styles['Normal'].alignment = 1
                lin[n] = item
                n += 1
            new_data.append(lin)
            m += 1

        tblstyle = TableStyle(
            [('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
             ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
             ('LEFTPADDING', (0, 0), (-1, -1), 5),
             ('RIGHTPADDING', (0, 0), (-1, -1), 5),
             ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
             ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])

        tbl = Table(new_data, colWidths=colwidths, repeatRows=1)

        tbl.setStyle(tblstyle)
        body.append(tbl)

        doc = SimpleDocTemplate(
            'tmp_rot_file.pdf', pagesize=landscape(letter),
            rightMargin=.5*inch, leftMargin=.5*inch, topMargin=.75*inch,
            bottomMargin=.5*inch)

        doc.build(body)

        pdf_old = open('tmp_rot_file.pdf', 'rb')
        pdf_reader = PyPDF4.PdfFileReader(pdf_old)
        pdf_writer = PyPDF4.PdfFileWriter()

        for pagenum in range(pdf_reader.numPages):
            page = pdf_reader.getPage(pagenum)
            page.rotateCounterClockwise(90)
            pdf_writer.addPage(page)

        pdf_out = open(self.filename, 'wb')
        pdf_writer.write(pdf_out)
        pdf_out.close()
        pdf_old.close()
        os.remove('tmp_rot_file.pdf')
