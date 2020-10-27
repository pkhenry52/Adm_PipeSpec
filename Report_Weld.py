from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Flowable, Paragraph, SimpleDocTemplate,\
     Spacer, Table, TableStyle, PageBreak
from reportlab.lib import colors
import PyPDF4
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

    def __init__(self, rptdata, Colnames, Colwdths, filename, ttl):
        # (table name, table data, table column names, table column
        # widths, name of PDF file)
        self.rptdata = rptdata
        self.columnames = Colnames
        self.colms = Colwdths
        self.filename = filename
        self.ttl = ttl

        self.width, self.height = letter
        self.textAdjust = 6.5

    def create_pdf(self):
        body = []

        styles = getSampleStyleSheet()
        spacer1 = Spacer(0, 0.25*inch)
        spacer2 = Spacer(0, 0.5*inch)

        tblstyle = TableStyle([
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])

        # provide the table name for the page header
        ptext = '<font size=14>%s</font>' % self.ttl
        body.append(Paragraph(ptext, styles["Heading2"]))

        line = LineDrw(500)

        if self.rptdata != []:
            txt = ('''These weld procedures are indicators only.<br/>Actual details
                    are needed from the WPS sheets.<br/>Either supplied by
                the company or supplied by<br/>the fabricator and approved
                by the company.''')
            ptxt = '<font size=12>%s</font>' % txt
            body.append(Paragraph(ptxt, styles["Normal"]))
            body.append(spacer1)

            body.append(line)
            body.append(spacer2)
            # separate the data into each procedure requirement
            for rowdata in self.rptdata:
                tbldata = []
                # add the column names to the table
                tbldata.append(self.columnames[1])
                n = 0
                for i in self.columnames[0]:
                    # add the requirement process type and notes
                    ptxt = ("<b>{0}:</b>      {1}<br/>".
                            format(i, rowdata[0][n]))
                    body.append(Paragraph(ptxt, styles["Normal"]))
                    body.append(spacer1)
                    n += 1

                for seg in rowdata[1:len(rowdata)]:
                    m = 0
                    seg = list(seg)
                    for item in seg:
                        # wrap any text which is longer than 10 characters
                        if type(item) == str:
                            if len(item) >= 10:
                                item = Paragraph(item, styles['Normal'])
                                seg[m] = item
                        m += 1
                    tbldata.append(tuple(seg))

                    colwdth1 = []
                    for i in self.colms[1]:
                        colwdth1.append(i * self.textAdjust)

                    tbl1 = Table(tbldata, colWidths=colwdth1)

                tbl1.setStyle(tblstyle)
                body.append(tbl1)
                body.append(spacer2)
                body.append(PageBreak())

        else:
            body.append(line)
            txt = 'have not been set up.'
            ptext = '<font size=14>%s</font>' % txt
            body.append(Paragraph(ptext, styles["Heading2"]))

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
