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

    def __init__(self, frmtitle, rptdata, colwdth, colnames,
                 rows_pg, colms_pg, filename, ttl=None):
        # (table name, table data, table column widths,
        # table column names, name of PDF file)
        self.rpttitle = frmtitle          # table name
        self.rptdata = rptdata                 # data to be printed
        self.columnames = colnames          # all the column titles
        self.colms = colwdth                # specified width of columns
        self.filename = filename            # name of the pdf file
        self.width, self.height = letter
        self.num_records = len(self.rptdata)
        self.ttl = ttl

        # adjust column width the correspond to width of text character
        self.textAdjust = 6.5
        # specified number of rows of data to be printed per page
        self.Rows_Pg = rows_pg
        # specified number of columns of data to be printed per page
        self.Colms_pg = colms_pg

    def tbldata(self, styles, start, finish, page):

        new_data = []
        rpt_data = []

        NoDataRow = page // 2 - 1
        if page % 2:
            NoDataRow = (page//2)

        for m in range((self.Rows_Pg*NoDataRow),
                       (self.Rows_Pg*NoDataRow + self.Rows_Pg)):
            if m > len(self.rptdata)-1:
                break
            element = self.rptdata[m]
            # odd number pages
            if start == 0:
                element = element[start:finish]
            # even number pages
            else:
                element = list(element[:1] + element[start:finish])
            rpt_data.append(element)

        # convert each string of data record from list to tuple
        rpt_data = [tuple(lin) for lin in rpt_data]

        # break long lines down into paragraph structure
        for lin in rpt_data:
            n = 0
            lin = list(lin)
            for item in lin:
                if type(item) == str:
                    if len(item) >= 20:
                        item = Paragraph(item, styles['Normal'])
                lin[n] = item
                n += 1
            new_data.append(lin)

        return new_data

    def create_pdf(self):

        body = []
        if not self.filename:
            exit()

        RptColNames = []
        RptColWdths = []

        styles = getSampleStyleSheet()
        spacer1 = Spacer(0, 0.25*inch)
        spacer2 = Spacer(0, 0.5*inch)

        # provide the table name for the page header
        if self.ttl is not None:
            txt = self.rpttitle + ' ' + self.ttl
        else:
            txt = self.rpttitle
        ptext = '<font size=14>%s</font>' % txt
        body.append(Paragraph(ptext, styles["Heading2"]))

        line = LineDrw(500)

        body.append(line)
        body.append(spacer1)

        tblstyle = TableStyle([
            ('INNERGRID', (0, 0), (-1, -1), 0.25, colors.black),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.black),
            ('LEFTPADDING', (0, 0), (-1, -1), 5),
            ('RIGHTPADDING', (0, 0), (-1, -1), 5),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])

        # Allow for specified number of rows of data per page this means
        # the number of pages is just the number or recordes divided by
        # number of rows_page
        if self.num_records % self.Rows_Pg == 0:
            NoPgs = (self.num_records//self.Rows_Pg) * 2
        else:
            NoPgs = (self.num_records//self.Rows_Pg + 1) * 2

        sect_pg = int(round(round(len(self.rptdata[0])/self.Colms_pg, 1)+.1))

        # build the report one page at a time based on even
        # and odd number pages starting at page 1
        for page in range(1, NoPgs+1, sect_pg):
            for m in range(1, sect_pg+1):
                start = (m-1) * self.Colms_pg
                finish = m * self.Colms_pg

                RptColNames = self.columnames[start:finish]
                RptColWdths = [i*self.textAdjust for
                               i in self.colms[start:finish]]

                if start != 0:
                    RptColNames.insert(0, self.columnames[0])
                    RptColWdths.insert(0, self.colms[0]*self.textAdjust)

                tbldata = self.tbldata(styles, start, finish, page)
                tbldata.insert(0, tuple(RptColNames))
                tbl = Table(tbldata, colWidths=RptColWdths)

                tbl.setStyle(tblstyle)
                body.append(spacer2)
                body.append(tbl)

            body.append(PageBreak())

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
