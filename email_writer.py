import dateutil.relativedelta
import win32com.client as win32
from datetime import datetime
import sys, os


class EmailWriter:

    def __init__(self, total_ga, potencijalni_linkovi, spojeno_ceka_esd, manje_15, vise_15, predikcija_ga, spojeno_ceka_esm, prethodni_mjesec):
        self.total_ga = total_ga
        self.potencijalni_linkovi = potencijalni_linkovi
        self.spojeno_ceka_esd = spojeno_ceka_esd
        self.manje_15 = manje_15
        self.vise_15 = vise_15
        self.predikcija_ga = predikcija_ga
        self.spojeno_ceka_esm = spojeno_ceka_esm
        self.prethodni_mjesec = prethodni_mjesec

    def DrugaPolovina(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.Subject = "ITS GA - Lokacija spojena, čeka se ESD | " + datetime.now().strftime('%d/%m/%Y')
        mail.To = "cpe_voditelji@A1.hr"
        mail.CC = "gordana.kovacevic@A1.hr ;alen.petric@A1.hr"
        mail.HTMLBody = fr"""
            Pozdrav, <br><br>
            Trenutna količina GAs za <b>{datetime.now().strftime('%m/%Y')}</b> je <b>{self.total_ga}</b>.<br><br>
            Trenutno u <b>ITS</b> imamo <b>{self.potencijalni_linkovi}</b> nova <b>potencijalna</b> linka, od čega je na:<br><br>
            <ul>
                <li><b>ESD-u: {self.spojeno_ceka_esd}</b></li>
                Od čega je u backlogu:<br>
                < 15 dana -> {self.manje_15} linkova<br>
                > 15 dana -> {self.vise_15} linkova<br><br>
                Od toga prema PowerBI report <b>predikciji</b> imamo <b>{self.predikcija_ga} potencijalna link(ov)a do kraja mjeseca.</b><br><br>
                Analiza po trenutnim statusima:<br>
                <img src= "{os.path.dirname(os.path.abspath(sys.argv[0]))}\usluge_cropped.png"><br>
                <li><b>ESM-u: {self.spojeno_ceka_esm}</b></li>
            </ul><br>
            Popis ESD ITS GA backloga nalazi se na linku: <span style="font-size:18"><a href="https://a1g.sharepoint.com/:x:/r/sites/o365ESD-ESMkoordinacija/Shared%20Documents/General/ITS%20GA,%20%C4%8Deka%20ESD.xlsx?d=wccefade660e746668e60a09a2ca476d6&csf=1&web=1&e=zQ9fbR">ITS GA, ČEKA ESD.xlsx</a></span><br><br>
            Lp,<br><br>
        """

        mail.Display()

    def PrvaPolovina(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mail.Subject = "ITS GA - Lokacija spojena, čeka se ESD | " + datetime.now().strftime('%d/%m/%Y')
        mail.To = "cpe_voditelji@A1.hr"
        mail.CC = "gordana.kovacevic@A1.hr ;alen.petric@A1.hr"
        mail.HTMLBody = fr"""
            Pozdrav, <br><br>
            Trenutna količina GAs za <b>{datetime.now().strftime('%m/%Y')}</b> je <b>{self.total_ga}</b>.<br><br>
            Trenutno u <b>ITS</b> imamo <b>{self.potencijalni_linkovi}</b> nova <b>potencijalna</b> linka, od čega je na:<br><br>
            <ul>
                <li><b>ESD-u: {self.spojeno_ceka_esd}</b></li>
                Od čega je u backlogu:<br>
                < 15 dana -> {self.manje_15} linkova<br>
                > 15 dana -> {self.vise_15} linkova<br><br>
                Od toga prema PowerBI report <b>predikciji</b> imamo <b>{self.predikcija_ga} potencijalna link(ov)a u sljedećih 15 dana.</b><br><br>
                Analiza po trenutnim statusima:<br>
                <img src= "{os.path.dirname(os.path.abspath(sys.argv[0]))}\usluge_cropped.png"><br>
                <li><b>ESM-u: {self.spojeno_ceka_esm}</b></li>
            </ul><br>
            Popis ESD ITS GA backloga nalazi se na linku: <span style="font-size:18"><a href="https://a1g.sharepoint.com/:x:/r/sites/o365ESD-ESMkoordinacija/Shared%20Documents/General/ITS%20GA,%20%C4%8Deka%20ESD.xlsx?d=wccefade660e746668e60a09a2ca476d6&csf=1&web=1&e=zQ9fbR">ITS GA, ČEKA ESD.xlsx</a></span><br><br>
            Lp,<br><br>
        """

        mail.Display()

    def PocetakMjeseca(self):
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)

        mjesecOduzimanje = dateutil.relativedelta.relativedelta(months=1)
        prosliMjesec = datetime.now() - mjesecOduzimanje
        prosliMjesec = datetime.strftime(prosliMjesec, "%m/%Y")

        mail.Subject = "ITS GA - Lokacija spojena, čeka se ESD | " + datetime.now().strftime('%d/%m/%Y')
        mail.To = "cpe_voditelji@A1.hr"
        mail.CC = "gordana.kovacevic@A1.hr ;alen.petric@A1.hr"
        mail.HTMLBody = fr"""
            Pozdrav, <br><br>
            Količina realiziranih GAs za <b>{prosliMjesec}</b> je <b>{self.prethodni_mjesec}</b>.<br>
            Trenutna količina GAs za <b>{datetime.now().strftime('%m/%Y')}</b> je <b>{self.total_ga}</b>.<br><br>
            Trenutno u <b>ITS</b> imamo <b>{self.potencijalni_linkovi}</b> nova <b>potencijalna</b> linka, od čega je na:<br><br>
            <ul>
                <li><b>ESD-u: {self.spojeno_ceka_esd}</b></li>
                Od čega je u backlogu:<br>
                < 15 dana -> {self.manje_15} linkova<br>
                > 15 dana -> {self.vise_15} linkova<br><br>
                Od toga prema PowerBI report <b>predikciji</b> imamo <b>{self.predikcija_ga} potencijalna link(ov)a u sljedećih 15 dana.</b><br><br>
                Analiza po trenutnim statusima:<br>
                <img src= "{os.path.dirname(os.path.abspath(sys.argv[0]))}\usluge_cropped.png"><br>
                <li><b>ESM-u: {self.spojeno_ceka_esm}</b></li>
            </ul><br>
            Popis ESD ITS GA backloga nalazi se na linku: <span style="font-size:18"><a href="https://a1g.sharepoint.com/:x:/r/sites/o365ESD-ESMkoordinacija/Shared%20Documents/General/ITS%20GA,%20%C4%8Deka%20ESD.xlsx?d=wccefade660e746668e60a09a2ca476d6&csf=1&web=1&e=zQ9fbR">ITS GA, ČEKA ESD.xlsx</a></span><br><br>
            Lp,<br><br>
        """

        mail.Display()
