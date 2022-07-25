from PyQt5 import QtWidgets, uic, QtGui, QtCore
from PyQt5.QtWidgets import QFileDialog, QMessageBox
import sys
import os
import shutil
import datetime
from zipfile import ZipFile
from program import *
import logging

logging.basicConfig(filename='app.log', filemode='w',
                    format='%(name)s - %(levelname)s - %(message)s')
logging.warning('This will get logged to a file')


class Ui(QtWidgets.QMainWindow):
    def __init__(self, deger, completed):
        super(Ui, self).__init__()
        uic.loadUi(os.getcwd()+'/gsv5.ui', self)  # Load the .ui file
        self.show()  # Show the GUI
        (self.nr, self.nc) = deger.shape
        self.deger = deger
        self.completed = completed
        self.yaz()
        self.ayar()
        app.processEvents()
        logging.warning('Program Başladı')
        
    def yaz(self):

        if self.sub_pc.isChecked():

            self.write_df_to_qtable(self.deger.iloc[:, 0:37],self.tablo,"alt")
        else:
            self.write_df_to_qtable(self.deger.iloc[:, [0,1,2,3,6,7,10,14,17,20,25,28,31,34]], self.tablo)

    def write_df_to_qtable(self, tablolar, tablo, tip="genel"):
        onceki = 0
        if type(tablolar) == pd.core.frame.DataFrame:
            tablolar = [tablolar]
        for df in tablolar:
            headers = list(df)
            tablo.setRowCount(df.shape[0])
            tablo.setColumnCount(df.shape[1])
            tablo.setHorizontalHeaderLabels(headers)
            # getting data from df is computationally costly so convert it to array first
            df_array = df.values
            if df.shape[0] == 0:
                row = 0
            for row in range(onceki, df.shape[0]):
                for col in range(df.shape[1]):
                 
                    # print(df_array[row, col])
                    if col > 2 and df_array[row, col] >= 100:
                        df_array[row, col] = 100

                    tablo.setItem(row, col, QtWidgets.QTableWidgetItem(
                    str(df_array[row, col])))
                    if tip=="alt":
                        font = QtGui.QFont()
                        font.setBold(True)
                        if col in [3,6,7,10,14,17,20,25,28,31,34] and df_array[row, col] >= 50:
                            tablo.item(row, col).setBackground(
                                QtGui.QColor(198, 255, 179))
                            tablo.item(row,col).setFont(font)
                        elif col in [3,6,7,10,14,17,20,25,28,31,34] and df_array[row, col] < 50:
                            tablo.item(row, col).setBackground(
                                QtGui.QColor(255, 128, 128))
                            tablo.item(row,col).setFont(font)
                    else:
                        if col >2 and df_array[row, col] >= 50:
                            tablo.item(row, col).setBackground(
                                QtGui.QColor(198, 255, 179))
                        elif col >2 and df_array[row, col] < 50:
                            tablo.item(row, col).setBackground(
                                QtGui.QColor(255, 128, 128))
                        
            for i in range(3, 37):
                tablo.setColumnWidth(i, 50)

            onceki = row+1
            logging.warning('Arayüz tablosuna eklendi')
    

    @staticmethod
    def uyarı(durum, mesajlar, report):
        logging.warning('!!!!Uyarı!!!!')
        mesaj = ""
        for i in mesajlar:
            mesaj = mesaj+i+os.linesep
        msg = QMessageBox()
        if durum == "Hata":
            msg.setIcon(QMessageBox.Critical)
            uyari = "İçeri aktarım gerçekleştirirken hata oluştu.Aşağıdaki dosyalar zaten içeri aktarılmış.\n"
            msg.setText(uyari)
            report.write(uyari)
        elif durum == "Tamamlananlar":
            msg.setIcon(QMessageBox.Information)
            uyari = "Aşağıdaki dosyalar başarı ile içeri aktarıldı.\n"
            msg.setText(uyari)
            report.write(uyari)
        elif durum == "Silindi":
            msg.setIcon(QMessageBox.Information)
            uyari = "Aşağıdaki dosyalar başarı ile silindi"
            msg.setText(uyari)
            report.write(uyari)
        elif durum == "Yok":
            msg.setIcon(QMessageBox.Warning)
            uyari = "Aşağıdaki dosya içeri aktarılanlar arasında bulunamadı"
            msg.setText(uyari)
            report.write(uyari)
        elif durum == "Bilinmeyen":
            msg.setIcon(QMessageBox.Question)
            uyari = "Aşağıdaki dosya içeri aktarılanlarılırken bilinmeyen hata oluştu ve içeri aktarılamadı"
            msg.setText(uyari)
            report.write(uyari)
        msg.setInformativeText(mesaj)
        report.write(mesaj)
        msg.setWindowTitle(durum)
        msg.exec_()
        # logging.warning(uyari)
        return report

    def son_aktarim_raporu(self):
        report = open(os.path.join(os.getcwd()+"/data/report.txt"), "r",encoding='utf-8')
        mesaj = report.read()
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("Son aktarım raporu aşağıdaki gibidir.")
        msg.setInformativeText(mesaj)
        msg.setWindowTitle("Son aktarım raporu")
        msg.exec_()
        logging.warning('Son aktarım rapor gösterildi')

    def sil(self):
        # pass
        # button = self.sender()
        # if button:
        #     row = self.tablo.indexAt(button.pos()).row()
        #     self.tablo.removeRow(row)
        filename, _ = QFileDialog.getOpenFileName(
                self, "İçeri aktarılacak dosyayı seçiniz")
        
            
        ret = QMessageBox.question(self, 'Seçilenleri Sil', "Seçili tüm kayıtlar silinecek, Emin misiniz?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            mezun=get_mezun(filename)
            a=self.deger["No"].isin(mezun["No"])
            self.deger.drop(self.deger["No"][a.tolist()].index,axis=0,inplace=True)
            store(self.deger, self.completed)
            self.yaz()
            QMessageBox.question(self, 'Seçilenleri Sil', "Ekteki dosyadaki öğrenciler silindi.",
                                   QMessageBox.Yes)
            logging.warning('Tüm veriler temizlendi')
        else:
            pass  
        
        # filename, _ = QFileDialog.getOpenFileName(
        #     self, "Silinecek dosyayı seçiniz")
        # report = open(os.path.join(os.getcwd()+"/data/report.txt"), "w")
        # list_of_student, completed = get_form(
        #     filename, self.completed, sil="Evet")
        # if type(list_of_student) != pd.core.frame.DataFrame:
        #     Ui.uyarı("Yok", [filename], report)
        #     return
        # pc = list(list_of_student.columns.values)[-2]
        # pc_ = pc + "+"
        # mask = (self.deger['No'].isin(list_of_student["No"]))
        # self.deger[mask][[pc, pc_]].replace(
        #     list_of_student[[pc, pc_]], inplace=True)

        # store(self.deger, self.completed)
        # self.write_df_to_qtable(self.deger.iloc[:, 0:14], self.tablo)
        # Ui.uyarı("Silindi", [filename.split("/")[-1]], report)
        
    def proje_picker(self,tip):
            if tip=="arastirma":
                title="Araştırma Projesi dosyalarının olduğu klasörü seçiniz."
            else:
                title="Tasarım Projesi dosyalarının olduğu klasörü seçiniz."
            folder = str(QFileDialog.getExistingDirectory(self, title))
            filenames = files(folder)
            report = open(os.path.join(os.getcwd()+"/data/report.txt"), "w",encoding='utf-8')
            hatalar = []
            aktarilanlar = []
            bilinmeyen = []
            for filename in filenames:
                if tip=="arastirma":
                    list_of_student, completed = get_arastirma_proje(
                        filename, self.completed)
                elif tip=="tasarim":
                    list_of_student, completed = get_tasarim_proje(
                        filename, self.completed)
                else:
                    return
                
                if type(list_of_student) != pd.core.frame.DataFrame:
                    if list_of_student == "zaten":
                        hatalar.append(os.path.basename(filename))
                    elif list_of_student == "bilinmeyen":
                        bilinmeyen.append(os.path.basename(filename))
                else:
                    # if:
                    #     self.deger = pd.concat([self.deger.iloc[:,[0,3:]], list_of_student]).groupby(
                    #     ['No']).sum().reset_index()
                 
                    # else:
                    self.deger = pd.concat([self.deger, list_of_student]).groupby(
                        ['No', 'Adı', 'Soyadı']).sum().reset_index()
                    aktarilanlar.append(os.path.basename(filename))
                
            store(self.deger, self.completed)
            self.yaz()
            if hatalar != []:
                report = Ui.uyarı("Hata", hatalar, report)
            if aktarilanlar != []:
                report = Ui.uyarı("Tamamlananlar", aktarilanlar, report)
            if bilinmeyen != []:
                report = Ui.uyarı("Bilinmeyen", bilinmeyen, report)
            report.close()
        
    def pick_new(self, secim, filetype="ziprar"):
        if secim == "folder":
            folder = str(QFileDialog.getExistingDirectory(
                self, "Ders dosyalarının olduğu klasörü seçiniz."))
       
            filenames = files(folder)
       
        else:
            filename, _ = QFileDialog.getOpenFileName(
                self, "İçeri aktarılacak dosyayı seçiniz")
            filenames = [filename]
        report = open(os.path.join(os.getcwd()+"/data/report.txt"), "w",encoding='utf-8')
        hatalar = []
        aktarilanlar = []
        bilinmeyen = []
        if filenames==None:
            return
        for filename in filenames:
            if filetype == "ziprar":
                if filename[-3:] == "zip":
                    list_of_student, completed = get_form(
                        filename, "zip", self.completed)
                elif filename[-3:] == "rar":
                    list_of_student, completed = get_form(
                        filename, "rar", self.completed)
                elif filename[-4:] == "xlsx" or filename[-3:] == "xls":
                    list_of_student, completed = get_form(
                        filename, "xlsx", self.completed)
                else:
                    return

            if type(list_of_student) != pd.core.frame.DataFrame:
                if list_of_student == "zaten":
                    hatalar.append(os.path.basename(filename))
                elif list_of_student == "bilinmeyen":
                    bilinmeyen.append(os.path.basename(filename))
            else:
                self.deger = pd.concat([self.deger, list_of_student]).groupby(
                    ['No', 'Adı', 'Soyadı']).sum().reset_index()
                aktarilanlar.append(os.path.basename(filename))
        store(self.deger, self.completed)
        self.yaz()
        if hatalar != []:
            report = Ui.uyarı("Hata", hatalar, report)
        if aktarilanlar != []:
            report = Ui.uyarı("Tamamlananlar", aktarilanlar, report)
        if bilinmeyen != []:
            report = Ui.uyarı("Bilinmeyen", bilinmeyen, report)
        report.close()
        # logging.warning("Aşağıdaki dosyalar programa eklenmek üzere seçildi"+"\n".join(filenames))


    def ayar(self):
        self.namesearch.textChanged.connect(self.namesearch_)
        self.nosearch.textChanged.connect(self.nosearch_)
        self.surnsearch.textChanged.connect(self.surnsearch_)
        self.import_2.clicked.connect(lambda: self.pick_new("folder"))
        self.D_import_2.clicked.connect(lambda: self.pick_new("folder"))
        self.import_tek.clicked.connect(lambda: self.pick_new("file"))
        self.delete_2.clicked.connect(self.sil)
        self.birlestir.clicked.connect(self.merge)
        self.last_report.clicked.connect(self.son_aktarim_raporu)
        self.check.clicked.connect(self.kontrol)
        self.D_check.clicked.connect(self.kontrol)
        self.D_proje_dersi.clicked.connect(lambda: self.proje_picker("arastirma"))
        self.proje_dersi.triggered.connect(lambda: self.proje_picker("arastirma"))
        self.D_tasarim_dersi.clicked.connect(lambda: self.proje_picker("tasarim"))
        self.tasarim_dersi.triggered.connect(lambda: self.proje_picker("tasarim"))
        self.D_mezun_sil.clicked.connect(self.sil)
        self.mezun_sil.triggered.connect(self.sil)
        self.delete_database.triggered.connect(self.hepsi_sil)
        self.yedekle.triggered.connect(self.yedek)
        self.import_yedek.triggered.connect(self.import_data)
        self.D_tum_export.triggered.connect(self.export_all)
        self.hakkinda.triggered.connect(self.yardim_)
        self.import_list.triggered.connect(self.derslistesi)
        self.sub_pc.stateChanged.connect(self.yaz)
        
        
    def derslistesi(self):

        self.aktarma = QtWidgets.QWidget()
        self.aktarma.setObjectName("aktarma")
        self.aktara = QtWidgets.QPlainTextEdit(self.aktarma)
        self.aktara.setGeometry(QtCore.QRect(10, 10, 221, 31))
        self.aktara.setObjectName("aktara")
        self.aktara.setPlaceholderText("Ara")
        self.akttable = QtWidgets.QTableWidget(self.aktarma)
        self.akttable.setGeometry(QtCore.QRect(10, 50, 891, 441))
        self.akttable.setRowCount(1)
        self.akttable.setColumnCount(6)
        self.akttable.setObjectName("akttable")
        self.aktkapat = QtWidgets.QPushButton(self.aktarma)
        self.aktkapat.setGeometry(QtCore.QRect(250, 10, 80, 31))
        self.aktkapat.setObjectName("aktkapat")
        self.donemlik.addTab(self.aktarma, "")
        self.aktkapat.setText("Kapat")
        self.donemlik.setTabText(
            self.donemlik.indexOf(self.aktarma), "Aktarma")
        self.aktara.textChanged.connect(self.dersara)
        self.aktkapat.clicked.connect(lambda: self.donemlik.removeTab(
            self.donemlik.indexOf(self.aktarma)))
        self.donemlik.setCurrentIndex(self.donemlik.indexOf(self.aktarma))
        derslist = []
        for i in self.completed:
            derslist.append(i.split("-", 3))
        self.df = pd.DataFrame(
            derslist, columns=['PÇ', 'Ders kodu', 'Ders Adı', 'Diğer'])
        self.write_df_to_qtable(self.df, self.akttable)

    def dersara(self):
        ara = self.aktara.toPlainText().replace("i", "İ").upper()
        matching = [
            s for s in self.completed if ara in s.replace("i", "İ").upper()]
        ind = []
        for x in matching:
            ind.append(self.completed.index(x))
        ind = np.unique(ind)
        self.write_df_to_qtable(
            self.df.iloc[ind, :], self.akttable)

    def hepsi_sil(self):

        ret = QMessageBox.question(self, 'Verileri Sil', "Eski tüm kayıtlar silinecek, Emin misiniz?",
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if ret == QMessageBox.Yes:
            shutil.rmtree(os.path.join(os.getcwd()+"/data"))
            self.deger, self.completed = database()
            self.yaz()
            logging.warning('Tüm veriler temizlendi')
        else:
            pass

    def yedek(self):
        name = QtWidgets.QFileDialog.getSaveFileName(self, 'Yedeği kaydet', "Yedek {}".format(
            datetime.date.today()), "", "Zip files (*.zip)")
        shutil.make_archive(name[0], 'zip', os.path.join(os.getcwd()+"/data"))
        logging.warning('Tüm veriler yedeklendi')
        self.deger.to_excel("liste.xlsx")

    def import_data(self):
        self.hepsi_sil()
        file, _ = QFileDialog.getOpenFileName(
            self, "İçeri aktarılacak yedeği seçiniz")
        zf = ZipFile(file)
        zf.extractall(os.path.join(os.getcwd()+"/data"))
        self.deger, self.completed = database()
        self.yaz()

    def yardim_(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        uyari = "Bu program Serdar Turgut İnce tarafından YTU Gemi İnşaatı bölümü müdek akreditasyonu çalışmaları kapsamında yapılmıştır"
        mesaj = "Programın kullanımı için YTU Beşiktaş Kampüsü T blok T-420 nolu adres dahili:3014 ile iletişime geçebilirsiniz "
        msg.setText(uyari)
        msg.setInformativeText(mesaj)
        msg.setWindowTitle("Hakkında")
        msg.exec_()

    def nosearch_(self):
        ara = self.nosearch.toPlainText().upper()
        matching1 = self.deger["No"].str.contains(ara)
        self.write_df_to_qtable(
            [self.deger[matching1].iloc[:, [0,1,2,3,6,9,12,15,18,19,20,21,22,23]]], self.tableWidget)

    def namesearch_(self):
        ara = self.namesearch.toPlainText().upper()
        matching1 = self.deger["Adı"].str.contains(ara)
        self.write_df_to_qtable(
            [self.deger[matching1].iloc[:, [0,1,2,3,6,9,12,15,18,19,20,21,22,23]]], self.tableWidget)

    def surnsearch_(self):
        ara = self.surnsearch.toPlainText().upper()
        matching1 = self.deger["Soyadı"].str.contains(ara)
        self.write_df_to_qtable(
            [self.deger[matching1].iloc[:, [0,1,2,3,6,9,12,15,18,19,20,21,22,23]]], self.tableWidget)

    def kontrol(self):
            title="Tez Almak isteyen öğrenci listelerinin yer aldığı klasörü seçiniz"
            folder = str(QFileDialog.getExistingDirectory(self, title))
            filenames = files(folder)
            ogrenciler=pd.DataFrame()
            ogrenciler["Grup Numarası"]=pd.NA
            for filename in filenames:
                try:
                    list_of_student= pd.read_excel(filename).iloc[:, [0]].set_axis(["No"], axis=1)
                
                
                    if type(list_of_student) != pd.core.frame.DataFrame:
                            bilinmeyen.append(os.path.basename(filename))
                    
                    for i in list_of_student.No:
                        matching=self.deger["No"].str.contains(i)
                        ogrenciler=pd.concat([ogrenciler,self.deger[matching].iloc[:, 0:14]]) 
                        ogrenciler["Grup Numarası"].fillna(os.path.basename(filename)[:-5],inplace=True)
                    
                except:
                    pass
            cols = ogrenciler.columns.tolist()
            cols = cols[1:]+cols[:1]
            ogrenciler = ogrenciler[cols]
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Information)
            uyari = "Tez alan öğrenciler kontrol edildi. Excel olarak program dizinine sonuçlar yazdırıldı \
                Ayrıca Arama ekranından ilgili öğrencileri kontrol edebilirsiniz."
            msg.setText(uyari)
            msg.setText(uyari)
            msg.setWindowTitle("Tez Kontrol")
            msg.exec_()
    
            ogrenciler.iloc[:, 3:14]=ogrenciler.iloc[:, 3:14].where(ogrenciler.iloc[:, 3:14]<=100,100)
    
            self.write_df_to_qtable(
            [ogrenciler.iloc[:, 0:14]], self.tableWidget)

            writer = pd.ExcelWriter('Tez Öğrencileri Çıktı Kontrolü.xlsx', engine='xlsxwriter')
            ogrenciler.to_excel(writer)
            workbook  = writer.book
            worksheet = writer.sheets['Sheet1']
            
            # Get the dimensions of the dataframe.
            (max_row, max_col) = ogrenciler.shape
            format2 = workbook.add_format({'bg_color': '#FFC7CE',
                               'font_color': '#9C0006'})

            # Add a format. Green fill with dark green text.
            format1 = workbook.add_format({'bg_color': '#C6EFCE',
                               'font_color': '#006100'})
            # Apply a conditional format to the required cell range.
            worksheet.conditional_format(f'E2:O{max_row+1}', {'type': 'cell',
                                                     'criteria': '>=',
                                                     'value': 50,
                                                     'format': format1})

            # Write another conditional format over the same range.
            worksheet.conditional_format(f'E2:O{max_row+1}', {'type': 'cell',
                                                            'criteria': '<',
                                                            'value': 50,
                                                            'format': format2})
            
            # Close the Pandas Excel writer and output the Excel file.
            writer.save()
            self.donemlik.setCurrentIndex(self.donemlik.indexOf(self.tab_4))

    def merge(self):
        self.deger["No"] = self.deger.apply(Ui.fix_number, axis=1)
        
        for i in self.deger["No"]:
            try:
                # self.deger=self.deger.reset_index()
                matching=self.deger["No"].str.contains(i)
                dublicate=self.deger[matching]
                if len(dublicate.index)>1:
                    first=dublicate.index[0]
                    second=dublicate.index[1]
                    self.deger.iloc[first,3:]=self.deger.iloc[first,3:].add(self.deger.iloc[second,3:],fill_value=0)
                    self.deger=self.deger.drop([second])
                    self.deger=self.deger.reset_index(drop=True)       
            except:
                print(f"şu öğrenci numarasında bir hata var {i}")
        store(self.deger, self.completed)
        self.yaz()
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        uyari = "Birleştirme başarıyla tamamlandı"
        msg.setText(uyari)
        msg.setWindowTitle("Birleştirme İşlemi")
        msg.exec_()
    
    @staticmethod
    def fix_number(row):
            if " " in row.No:
                return row.No.replace(" ","")
            try:
                int(row.No[0])
                return row.No
            except:
                for i,j in enumerate(row.No):
                    if j.isdigit():
                        return row.No[i:]
                    
    def export_all(self):
        ogrenciler= self.deger.iloc[:,0:14]
        ogrenciler.iloc[:, 3:14]=ogrenciler.iloc[:, 3:14].where(ogrenciler.iloc[:, 3:14]<=100,100)

        writer = pd.ExcelWriter('Tüm Öğrenciler.xlsx', engine='xlsxwriter')
        ogrenciler.to_excel(writer)
        workbook  = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Get the dimensions of the dataframe.
        (max_row, max_col) = ogrenciler.shape
        format2 = workbook.add_format({'bg_color': '#FFC7CE',
                           'font_color': '#9C0006'})

        # Add a format. Green fill with dark green text.
        format1 = workbook.add_format({'bg_color': '#C6EFCE',
                           'font_color': '#006100'})
        # Apply a conditional format to the required cell range.
        worksheet.conditional_format(f'E2:O{max_row+1}', {'type': 'cell',
                                                 'criteria': '>=',
                                                 'value': 50,
                                                 'format': format1})

        # Write another conditional format over the same range.
        worksheet.conditional_format(f'E2:O{max_row+1}', {'type': 'cell',
                                                        'criteria': '<',
                                                        'value': 50,
                                                        'format': format2})
        
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()          
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        uyari = "Tüm öğrencilerin listesi excel olarak program dizinine sonuçlar yazdırıldı."
        msg.setText(uyari)
        msg.setText(uyari)
        msg.setWindowTitle("Aktarma Tamamlandı")
        msg.exec_()
    
        
        


tum_ogrenciler, completed = database()

# Create an instance of QtWidgets.QApplication
app = QtWidgets.QApplication(sys.argv)
window = Ui(tum_ogrenciler, completed)  # Create an instance of our class
app.exec_()  # Start the application
