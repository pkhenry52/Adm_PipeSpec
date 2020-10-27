from reportlab.lib.pagesizes import letter
from reportlab.pdfgen.canvas import Canvas
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import Flowable, Paragraph, Frame, Spacer
from reportlab.lib.enums import TA_LEFT, TA_RIGHT


class LineDrw(Flowable):

    def __init__(self, width, height=0):
        Flowable.__init__(self)
        self.width = width
        self.height = height

    def __repr__(self):
        return "Line(w=%s)" % self.width

    def draw(self):
        self.canv.line(0, self.height, self.width, self.height)


class HTRForm:
    def __init__(self, pdf_file, t_data, ta_data, chk_data=None):
        self.t = t_data
        self.ta = ta_data
        self.chk = chk_data
        self.pdf_file = pdf_file

    def create(self):
        styles = getSampleStyleSheet()

        aW, aH = letter

        nrm = styles['Normal']
        hd1 = styles['Heading1']
        hd2 = styles['Heading2']

        styles.add(ParagraphStyle(name='Right', fontSize=12, spaceAfter=15,
                                  alignment=TA_RIGHT))
        styles.add(ParagraphStyle(name='Left', fontSize=12, spaceAfter=15,
                                  borderColor='black', borderRadius=3,
                                  borderPadding=5, borderWidth=1,
                                  alignment=TA_LEFT))

        basepg = Canvas(self.pdf_file)

        # Form title
        ttl = []
        ttl.append(Paragraph('Hydro Test Report', hd1))
        ttl.append(LineDrw(500))
        frm = Frame(15*mm, 260*mm, width=150*mm, height=25*mm, showBoundary=0)
        frm.addFromList(ttl, basepg)

        # First text area box
        dscrp = []
        dscrp.append(Paragraph('Project Description', hd2))
        dscrp.append(Paragraph(self.ta['ta_1'], nrm))
        frm1 = Frame(15*mm, 240*mm, width=180*mm, height=25*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # Section of labels for input text
        info = []
        info.append(Paragraph('<b><i>Work Order:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Project Engineer:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Contact Information:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Fabricator:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Date Of Issue:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Pipe Specification Code:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>P&ID Line Number:</i></b>',
                              style=styles['Right']))
        frmLft = Frame(15*mm, 140*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmLft.addFromList(info, basepg)

        # list of first section of input text boxes
        inpt = []
        for n in range(1, 8):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
        frmRgt = Frame(80*mm, 140*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmRgt.addFromList(inpt, basepg)

        # second textarea box
        dscrp = []
        dscrp.append(Paragraph('Related Drawings and Documents:', hd2))
        dscrp.append(Paragraph(self.ta['ta_2'], nrm))
        frm1 = Frame(15*mm, 80*mm, width=180*mm, height=65*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        basepg.showPage()

        # title for beginning of operating conditions
        dscrp = []
        dscrp.append(LineDrw(500))
        dscrp.append(Paragraph('Operating Conditions', hd2))
        dscrp.append(Spacer(1, 13))
        frm2 = Frame(15*mm, 260*mm, width=180*mm, height=15*mm, showBoundary=0)
        frm2.addFromList(dscrp, basepg)

        # text box labels for operating conditions
        dscrp = []
        dscrp.append(Paragraph('<b><i>Design Pressure (psig):</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 5))
        dscrp.append(Paragraph('<b><i>Hydro Test Pressure (psig):</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 5))
        dscrp.append(Paragraph('<b><i>Hydro Test Temperature (F):</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 7))
        dscrp.append(Paragraph('<b><i>Hydro Test Media:</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 6))
        dscrp.append(Paragraph('<b><i>Test Start Time (hr:min):</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 6))
        dscrp.append(Paragraph('<b><i>Test Sign Off Time (hr:min):</i></b>',
                               style=styles['Right']))
        dscrp.append(Spacer(1, 6))
        dscrp.append(Paragraph('<b><i>Test Duration (min):</i></b>',
                               style=styles['Right']))
        frmLft = Frame(15*mm, 180*mm, width=75*mm, height=85*mm,
                       showBoundary=0)
        frmLft.addFromList(dscrp, basepg)

        # input text boxes for hydro test conditions
        inpt = []
        for n in range(8, 15):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 7))
        frmRgt = Frame(95*mm, 180*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmRgt.addFromList(inpt, basepg)

        # title for beginning of operating conditions
        dscrp = []
        dscrp.append(LineDrw(500))
        dscrp.append(Paragraph('Testing Gauges Record', hd2))
        dscrp.append(Spacer(1, 13))
        frm2 = Frame(15*mm, 165*mm, width=180*mm, height=15*mm, showBoundary=0)
        frm2.addFromList(dscrp, basepg)

        # text box labels for gage info
        info = []
        info.append(Paragraph('<b><i>Gauge 1 ID:</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 8))
        info.append(Paragraph('<b><i>Range:</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 8))
        info.append(Paragraph('<b><i>Gauge 2 ID:</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 8))
        info.append(Paragraph('<b><i>Range:</i></b>',
                    style=styles['Right']))
        frmLft1 = Frame(10*mm, 110*mm, width=35*mm, height=55*mm,
                        showBoundary=0)
        frmLft1.addFromList(info, basepg)

        # input text data for gage info
        inpt = []
        for n in range(15, 19):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 8))
        frmLft2 = Frame(50*mm, 110*mm, width=45*mm, height=55*mm,
                        showBoundary=0)
        frmLft2.addFromList(inpt, basepg)

        # text box labels for gage info
        info = []
        info.append(Paragraph('<b><i>Calibration Date:</i></b>', style=styles['Right']))
        info.append(Spacer(1, 28))
        info.append(Paragraph('<b><i>Calibration Date:</i></b>', style=styles['Right']))
        frmRgt1 = Frame(100*mm, 117*mm, width=35*mm, height=50*mm,
                        showBoundary=0)
        frmRgt1.addFromList(info, basepg)

        # text box input data for gage info
        inpt = []
        for n in range(19, 21):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 40))
        frmRgt2 = Frame(140*mm, 115*mm, width=35*mm, height=50*mm,
                        showBoundary=0)
        frmRgt2.addFromList(inpt, basepg)

        # second textarea box
        dscrp = []
        dscrp.append(Paragraph('Comments:', hd2))
        dscrp.append(Paragraph(self.ta['ta_3'], nrm))
        frm1 = Frame(15*mm, 60*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # sign off text boxes
        info = []
        info.append(Paragraph('<b><i>Owner Inspector:</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>QC Inspector:</i></b>',
                    style=styles['Right']))
        frmLft1 = Frame(10*mm, 20*mm, width=45*mm, height=30*mm,
                        showBoundary=0)
        frmLft1.addFromList(info, basepg)

        # input text data
        inpt = []
        for n in range(21, 23):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmLft2 = Frame(60*mm, 20*mm, width=45*mm, height=30*mm,
                        showBoundary=0)
        frmLft2.addFromList(inpt, basepg)

        # text box labels
        info = []
        for n in range(2):
            info.append(Paragraph('<b><i>Date:</i></b>', style=styles['Right']))
            info.append(Spacer(1, 13))
        frmRgt1 = Frame(120*mm, 20*mm, width=25*mm, height=30*mm,
                        showBoundary=0)
        frmRgt1.addFromList(info, basepg)

        # text box input data
        inpt = []
        for n in range(23, 25):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmRgt2 = Frame(150*mm, 20*mm, width=25*mm, height=30*mm,
                        showBoundary=0)
        frmRgt2.addFromList(inpt, basepg)

        basepg.save()
