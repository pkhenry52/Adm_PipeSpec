
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


class RPSForm:
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
        ttl.append(Paragraph('Request for New Pipe Specification', hd1))
        ttl.append(LineDrw(500))
        frm = Frame(15*mm, 260*mm, width=150*mm, height=25*mm, showBoundary=0)
        frm.addFromList(ttl, basepg)

        # First text area box
        dscrp = []
        txt = '''This form is to be used when a new commodity property has been approved and needs a related piping specification.
        Do not submit this form until the related commodity property has been approved by operations and/or process engineering.'''
        dscrp.append(Paragraph(txt, nrm))
        frm1 = Frame(15*mm, 240*mm, width=180*mm, height=25*mm, showBoundary=0)
        frm1.addFromList(dscrp, basepg)

        # Section of labels for input text
        info = []
        info.append(Paragraph('<b><i>Project Engineer:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Contact Information:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>Date Of Issue:</i></b>',
                              style=styles['Right']))
        info.append(Paragraph('<b><i>P&ID Line Number:</i></b>',
                              style=styles['Right']))
        frmLft = Frame(15*mm, 150*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmLft.addFromList(info, basepg)

        # list of first section of input text boxes
        inpt = []
        for n in range(1, 5):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
        frmRgt = Frame(80*mm, 150*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmRgt.addFromList(inpt, basepg)

        # second textarea box
        dscrp = []
        dscrp.append(Paragraph('Related Drawings or Documentation:', hd2))
        dscrp.append(Paragraph(self.ta['ta_1'], nrm))
        frm1 = Frame(15*mm, 140*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # title for beginning of operating conditions
        dscrp = []
        dscrp.append(LineDrw(500))
        dscrp.append(Paragraph('Operating Conditions', hd2))
        dscrp.append(Spacer(1, 13))
        frm2 = Frame(15*mm, 120*mm, width=180*mm, height=15*mm, showBoundary=0)
        frm2.addFromList(dscrp, basepg)

        # text box labels for operating conditions
        dscrp = []
        dscrp.append(Paragraph('<b><i>Similar Piping Specification:</i></b>',
                               style=styles['Right']))
        dscrp.append(Paragraph('<b><i>Maximum Operating Pressure (psig):</i></b>',
                               style=styles['Right']))
        dscrp.append(Paragraph('<b><i>Maximum Operating Temperature (F):</i></b>',
                               style=styles['Right']))
        dscrp.append(Paragraph('<b><i>Maximum Design Pressure (psig):</i></b>',
                               style=styles['Right']))
        dscrp.append(Paragraph('<b><i>Maximum Design Temperature (F):</i></b>',
                               style=styles['Right']))
        dscrp.append(Paragraph('<b><i>Minimum Design Temperature (F):</i></b>',
                               style=styles['Right']))
        frmLft = Frame(15*mm, 35*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmLft.addFromList(dscrp, basepg)

        # input text boxes for operating conditions
        inpt = []
        for n in range(5, 11):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 10))
        frmRgt = Frame(80*mm, 35*mm, width=65*mm, height=85*mm,
                       showBoundary=0)
        frmRgt.addFromList(inpt, basepg)

        basepg.showPage()

        # fourth textarea box
        dscrp = []
        dscrp.append(Paragraph('Description of Fluid', hd2))
        dscrp.append(Paragraph(self.ta['ta_2'], nrm))
        frm1 = Frame(15*mm, 240*mm, width=180*mm, height=45*mm, showBoundary=1)
        frm1.addFromList(dscrp, basepg)

        # text box labels
        info = []
        info.append(Paragraph('<b><i>Approved Pipe Specification</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>QC Manager Approval</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>Operations Approval</i></b>',
                    style=styles['Right']))
        info.append(Spacer(1, 13))
        info.append(Paragraph('<b><i>Project Manager Approval</i></b>',
                    style=styles['Right']))
        frmLft1 = Frame(10*mm, 140*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmLft1.addFromList(info, basepg)

        # input text data
        inpt = []
        for n in range(11, 15):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmLft2 = Frame(80*mm, 140*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmLft2.addFromList(inpt, basepg)

        # text box labels
        info = []
        for n in range(4):
            info.append(Paragraph('<b><i>Date</i></b>', style=styles['Right']))
            info.append(Spacer(1, 13))
        frmRgt1 = Frame(95*mm, 140*mm, width=65*mm, height=65*mm,
                        showBoundary=0)
        frmRgt1.addFromList(info, basepg)

        # text box input data
        inpt = []
        for n in range(15, 19):
            indx = 't_' + str(n)
            inpt.append(Paragraph(self.t[indx], style=styles['Left']))
            inpt.append(Spacer(1, 13))
        frmRgt2 = Frame(160*mm, 140*mm, width=35*mm, height=65*mm,
                        showBoundary=0)
        frmRgt2.addFromList(inpt, basepg)

        basepg.save()
