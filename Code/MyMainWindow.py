# -*- coding:GBK -*-

import sys
from PyQt5.QtWidgets import  (QApplication, QMainWindow, QFileDialog)
from PyQt5.QtCore import  pyqtSlot,QDir
import pandas as pd
import os
from docx import Document
import xlrd
from win32com.client import Dispatch,constants,gencache
import pdfplumber
from smtplib import SMTP_SSL
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

from ui_MainWindow import Ui_MainWindow
class QmyMainWindow(QMainWindow):

    def __init__(self, parent=None):
        super().__init__(parent)   #调用父类构造函数，创建窗体
        self.ui=Ui_MainWindow()    #创建UI对象
        self.ui.setupUi(self)      #构造UI界面

        self.setWindowTitle("Office Automation Tool")
        self.setCentralWidget(self.ui.tabWidget)



#====================Excel板块====================
#====================多表合并====================
    #多表合并功能
    def __ExcelMerge(self,excellist='',saveexcel=''):
        # 文件路径
        file_dir = excellist
        # 构建新的表格名称
        new_filename = saveexcel
        # 找到文件路径下的所有表格名称，返回列表
        file_list = os.listdir(file_dir)
        new_list = []

        for file in file_list:
            # 重构文件路径
            file_path = os.path.join(file_dir, file)
            # 将excel转换成DataFrame
            dataframe = pd.read_excel(file_path)
            # 保存到新列表中
            new_list.append(dataframe)

        # 多个DataFrame合并为一个
        df = pd.concat(new_list)
        # 写入到一个新excel表中
        df.to_excel(new_filename, index=False)

    @pyqtSlot() #选择要合并的文件夹
    def on_btn_MergeExcelList_clicked(self):
        curPath=QDir.currentPath()  #获取系统当前目录
        dlgTitle="选择目录"
        selectedDir=QFileDialog.getExistingDirectory(self,
                    dlgTitle,curPath,QFileDialog.ShowDirsOnly)
        self.ui.lineEditMergeExcelList.setText(selectedDir)

    @pyqtSlot() #选择保存路径
    def on_btn_SaveMergeExcel_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '保存文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_SaveMergeExcel.setText(filename)

    @pyqtSlot() #一键合并
    def on_btn_ExcelMerge_clicked(self):
        try:
            self.ui.textEdit_ExcelStatus.clear()
            excellist=self.ui.lineEditMergeExcelList.text()
            saveexcel=self.ui.lineEdit_SaveMergeExcel.text()
            self.__ExcelMerge(excellist=excellist,saveexcel=saveexcel)
            self.ui.textEdit_ExcelStatus.setPlainText("合并成功！")
            self.ui.lineEditMergeExcelList.clear()
            self.ui.lineEdit_SaveMergeExcel.clear()
        except:
            self.ui.textEdit_ExcelStatus.clear()
            self.ui.textEdit_ExcelStatus.setPlainText("合并失败！请检查是否操作正确！")
            self.ui.lineEditMergeExcelList.clear()
            self.ui.lineEdit_SaveMergeExcel.clear()

#====================消除重复====================
    #消除重复功能
    def __ExcelDuplication(self,Duplicationexcel='',
                           saveexcel='',
                           header=''):
        excel = pd.read_excel(Duplicationexcel)
        excel.drop_duplicates(subset=header, inplace=True)
        excel.to_excel(saveexcel)

    @pyqtSlot() #打开有重复数据的excel文件
    def on_btn_DuplicationExcel_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEditDuplicationExcel.setText(filename)

    @pyqtSlot() #保存消除重复后的文件
    def on_btn_SaveDuplicationExcel_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '保存文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_SaveDuplicationExcel.setText(filename)

    @pyqtSlot() #一键消除
    def on_btn_ExcelDuplication_clicked(self):
        try:
            self.ui.textEdit_ExcelStatus.clear()
            Duplicationexcel=self.ui.lineEditDuplicationExcel.text()
            saveexcel=self.ui.lineEdit_SaveDuplicationExcel.text()
            header=self.ui.lineEdit_ExcelDuplicationCol.text()
            self.__ExcelDuplication(Duplicationexcel=Duplicationexcel,
                                    saveexcel=saveexcel,header=header)
            self.ui.textEdit_ExcelStatus.setPlainText("消除成功！")
            self.ui.lineEditDuplicationExcel.clear()
            self.ui.lineEdit_SaveDuplicationExcel.clear()
            self.ui.lineEdit_ExcelDuplicationCol.clear()
        except:
            self.ui.textEdit_ExcelStatus.clear()
            self.ui.textEdit_ExcelStatus.setPlainText("消除失败！请检查是否操作正确！")
            self.ui.lineEditDuplicationExcel.clear()
            self.ui.lineEdit_SaveDuplicationExcel.clear()
            self.ui.lineEdit_ExcelDuplicationCol.clear()

#====================自动旋转====================
    def __ExcelRotate(self,excelfile='',sheetname='',savefile=''):
        excel = pd.read_excel(excelfile, sheet_name=sheetname, dtype=str)
        table = excel.transpose()
        table.to_excel(savefile)

    @pyqtSlot() #打开所要旋转的文件
    def on_btn_ExcelRotateFile_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_ExcelRotateFile.setText(filename)

    @pyqtSlot()  # 保存旋转后的文件
    def on_btn_SaveExcelRotate_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '保存文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_SaveExcelRotate.setText(filename)

    @pyqtSlot() #一键旋转
    def on_btn_ExcelRotate_clicked(self):
        try:
            self.ui.textEdit_ExcelStatus.clear()
            excelfile=self.ui.lineEdit_ExcelRotateFile.text()
            sheetname=self.ui.lineEdit_ExcelRotatesheet.text()
            savefile=self.ui.lineEdit_SaveExcelRotate.text()
            self.__ExcelRotate(excelfile=excelfile,sheetname=sheetname,savefile=savefile)
            self.ui.textEdit_ExcelStatus.setPlainText("旋转成功！")
            self.ui.lineEdit_ExcelRotateFile.clear()
            self.ui.lineEdit_SaveExcelRotate.clear()
            self.ui.lineEdit_ExcelRotatesheet.clear()
        except:
            self.ui.textEdit_ExcelStatus.clear()
            self.ui.textEdit_ExcelStatus.setPlainText("旋转失败！请检查操作是否正确！")
            self.ui.lineEdit_ExcelRotateFile.clear()
            self.ui.lineEdit_SaveExcelRotate.clear()
            self.ui.lineEdit_ExcelRotatesheet.clear()




#====================Word板块====================
#====================模板套用====================
    #新旧文字转换
    def change_text(self,document='',old_text='', new_text=''):
        document=Document(document)
        all_paragraphs = document.paragraphs
        for paragraph in all_paragraphs:
            for run in paragraph.runs:
                run_text = run.text.replace(old_text, new_text)
                run.text = run_text
        all_tables = document.tables
        for table in all_tables:
            for row in table.rows:
                for cell in row.cells:
                    cell_text = cell.text.replace(old_text, new_text)
                    cell.text = cell_text

    #模板套用功能
    def __WordModuleUse(self,wordmodule='',excelmodule='',savelist=''):
        xlsx = xlrd.open_workbook(excelmodule)
        sheet = xlsx.sheet_by_index(0)

        for table_row in range(1, sheet.nrows):
            document = Document(wordmodule)
            for table_col in range(0, sheet.ncols):
                self.change_text(wordmodule, str(sheet.cell_value(0, table_col)), str(sheet.cell_value(table_row, table_col)))

            document.save(savelist + '/' + '%s.docx' % str(sheet.cell_value(table_row, 0)))
            self.ui.textEdit_WordStatus.setText('%s完成' % str(sheet.cell_value(table_row, 0)))

    @pyqtSlot() #打开Word模板
    def on_btn_WordModule_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'Microsoft Word 文档(*.docx)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_WordModule.setText(filename)

    @pyqtSlot()  # 打开有模板数据的excel文件
    def on_btn_WordModuleExcel_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'Microsoft Excel 工作表(*.xlsx)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_WordModuleExcel.setText(filename)

    @pyqtSlot()  # 选择要保存模板的文件夹
    def on_btn_SaveWordModule_clicked(self):
        curPath = QDir.currentPath()  # 获取系统当前目录
        dlgTitle = "选择目录"
        selectedDir = QFileDialog.getExistingDirectory(self,
                                                       dlgTitle, curPath, QFileDialog.ShowDirsOnly)
        self.ui.lineEdit_SaveWordModule.setText(selectedDir)

    @pyqtSlot()  # 转换Word模板
    def on_btn_WordModuleTransform_clicked(self):
        try:
            self.ui.textEdit_WordStatus.clear()
            wordmodule=self.ui.lineEdit_WordModule.text()
            excelmodule=self.ui.lineEdit_WordModuleExcel.text()
            savelist=self.ui.lineEdit_SaveWordModule.text()
            self.__WordModuleUse(wordmodule=wordmodule,excelmodule=excelmodule,savelist=savelist)
            self.ui.textEdit_WordStatus.setPlainText("转换成功！")
            self.ui.lineEdit_WordModule.clear()
            self.ui.lineEdit_WordModuleExcel.clear()
            self.ui.lineEdit_SaveWordModule.clear()
        except:
            self.ui.textEdit_WordStatus.clear()
            self.ui.textEdit_WordStatus.setPlainText("转换失败！请检查是否操作正确！")
            self.ui.lineEdit_WordModule.clear()
            self.ui.lineEdit_WordModuleExcel.clear()
            self.ui.lineEdit_SaveWordModule.clear()

#====================Word转PDF====================
#====================转换====================
    def __WordToPDF(self,docxpath='',pdfpath=''):
        docx_path = docxpath
        pdf_path = pdfpath

        gencache.EnsureModule('{00020905-0000-0000-C000-000000000046}', 0, 8, 4)

        wd = Dispatch("Word.Application")

        doc = wd.Documents.Open(docx_path, ReadOnly=1)
        doc.ExportAsFixedFormat(pdf_path, constants.wdExportFormatPDF, Item=constants.wdExportDocumentWithMarkup,
                                CreateBookmarks=constants.wdExportCreateHeadingBookmarks)

        wd.Quit(constants.wdDoNotSaveChanges)

    @pyqtSlot() #选择所要转换的word文件
    def on_btn_transWord_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'Microsoft Word 文档(*.docx)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_transWord.setText(filename)

    @pyqtSlot()  # 保存转换后的PDF文件
    def on_btn_transPDF_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '保存文件'
        filt = 'PDF文件(*.pdf)'
        filename, filtUsed = QFileDialog.getSaveFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_transPDF.setText(filename)

    @pyqtSlot() #一键转换
    def on_btn_WordtoPDF_clicked(self):
        try:
            self.ui.textEdit_WordStatus.clear()
            docxpath=self.ui.lineEdit_transWord.text()
            pdfpath=self.ui.lineEdit_transPDF.text()
            self.__WordToPDF(docxpath=docxpath,pdfpath=pdfpath)
            self.ui.textEdit_WordStatus.setPlainText("转换成功！")
            self.ui.lineEdit_transWord.clear()
            self.ui.lineEdit_transPDF.clear()
        except:
            self.ui.textEdit_WordStatus.clear()
            self.ui.textEdit_WordStatus.setPlainText("转换失败！请检查是否操作正确！")
            self.ui.lineEdit_transWord.clear()
            self.ui.lineEdit_transPDF.clear()




#====================PDF板块====================
#====================PDF文字提取====================
    #==========文字单页提取功能==========
    def __PDFtextpage(self,filename='',page=1):
        with pdfplumber.open(filename) as pdf:
            text_page=pdf.pages[page-1]
            text=text_page.extract_text()
            self.ui.textEdit_PDF.setPlainText(text)

    #==========提取PDF全部文字
    def __PDFtextpages(self,filename=''):
        with pdfplumber.open(filename) as pdf:
            for page in pdf.pages:
                text=page.extract_text()
            self.ui.textEdit_PDF.setPlainText(text)

    @pyqtSlot() #打开PDF文件
    def on_btn_PDFTextfile_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = 'PDF 文件(*.pdf)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_PDFTextfile.setText(filename)

    @pyqtSlot() #提取该页
    def on_btn_PDFTextgetpage_clicked(self):
        try:
            self.ui.textEdit_PDF.clear()
            self.ui.textEdit_PDFStatus.clear()
            filename=self.ui.lineEdit_PDFTextfile.text()
            page=self.ui.spinBox_PDFTextpage.value()
            self.__PDFtextpage(filename=filename,page=page)
            self.ui.textEdit_PDFStatus.setPlainText("提取成功！")
        except:
            self.ui.textEdit_PDF.clear()
            self.ui.textEdit_PDFStatus.clear()
            self.ui.textEdit_PDFStatus.setPlainText("提取失败！请检查是否操作正确！")

    @pyqtSlot()  # 全部提取
    def on_btn_PDFTextgetpages_clicked(self):
        try:
            self.ui.textEdit_PDF.clear()
            self.ui.textEdit_PDFStatus.clear()
            filename = self.ui.lineEdit_PDFTextfile.text()
            self.__PDFtextpages(filename=filename)
            self.ui.textEdit_PDFStatus.setPlainText("提取成功！")
        except:
            self.ui.textEdit_PDF.clear()
            self.ui.textEdit_PDFStatus.clear()
            self.ui.textEdit_PDFStatus.setPlainText("提取失败！请检查是否操作正确！")

#====================邮件板块====================
#====================文本邮件====================
    #发送文本邮件功能
    def __TextEmail(self,host_server='',sender_email='',password='',
                    receiver='',main_title='',main_content=''):
        host_server = host_server  # 邮箱SMTP服务器
        sender_email = sender_email  # 发件人邮箱
        password = password  # 邮箱密码

        sender_qq_email = sender_email  # 发件人邮箱
        receiver = receiver  # 收件人邮箱

        main_title = main_title  # 邮件标题
        # 邮件正文
        main_content = main_content

        msg = MIMEMultipart()  # 邮件主体
        msg['Subject'] = Header(main_title, 'utf-8')
        msg['From'] = sender_qq_email
        msg['To'] = Header('Test', 'utf-8')
        msg.attach(MIMEText(main_content, 'plain', 'utf-8'))

        smtp = SMTP_SSL(host_server)  # ssl登陆
        smtp.login(sender_email, password)
        smtp.sendmail(sender_qq_email, receiver, msg.as_string())
        smtp.quit()

    @pyqtSlot() #清除文本内容
    def on_btn_TextEmailclearcontent_clicked(self):
        self.ui.textEdit_emailContent.clear()

    @pyqtSlot() #一键发送
    def on_btn_TextEmailSend_clicked(self):
        try:
            self.ui.textEdit_emailstatus.clear()
            host_server=self.ui.comboBox_TextEmailsmtpserver.currentText()
            sender_email=self.ui.lineEdit_TextEmailsenderemail.text()
            password=self.ui.lineEdit_TextEmailsenderpassword.text()
            receiver=self.ui.lineEdit_TextEmailreceiveremail.text()
            main_title=self.ui.lineEdit_TextEmailtitle.text()
            main_content=str(self.ui.textEdit_emailContent.toPlainText())
            self.__TextEmail(host_server=host_server,sender_email=sender_email,
                            password=password,receiver=receiver,
                            main_title=main_title,main_content=main_content)
            self.ui.textEdit_emailstatus.setPlainText("发送成功！")
        except:
            self.ui.textEdit_emailstatus.clear()
            self.ui.textEdit_emailstatus.setPlainText("发送失败！请检查是否操作正确！")

    #带附件邮件
    def __FileEmail(self,host_server='',sender_email='',password='',
                    receiver='',main_title='',main_content='',filename=''):
        host_server = host_server  # qq邮箱SMTP服务器
        sender_email = sender_email  # 发件人邮箱
        password = password  # 邮箱授权码

        sender_qq_email = sender_email  # 发件人邮箱
        receiver = receiver  # 收件人邮箱

        main_title = main_title  # 邮件标题
        # 邮件正文
        main_content = main_content

        msg = MIMEMultipart()  # 邮件主体
        msg['Subject'] = Header(main_title, 'utf-8')
        msg['From'] = sender_qq_email
        msg['To'] = Header('Test', 'utf-8')

        msg.attach(MIMEText(main_content, 'plain', 'utf-8'))

        attachment = MIMEApplication(open(filename, 'rb').read())
        new_filename=filename.split("/")[-1]
        attachment.add_header('Content-Disposition', 'attachment', filename=new_filename)

        msg.attach(attachment)

        smtp = SMTP_SSL(host_server)  # ssl登陆
        smtp.set_debuglevel(0)
        smtp.ehlo(host_server)
        smtp.login(sender_email, password)
        smtp.sendmail(sender_qq_email, receiver, msg.as_string())
        smtp.quit()

    @pyqtSlot() #清除文本内容
    def on_btn_FileEmailclearcontent_clicked(self):
        self.ui.textEdit_emailContent.clear()

    @pyqtSlot() #获取附件
    def on_btn_FindFileEmail_clicked(self):
        curPath = QDir.currentPath()
        dlgTitle = '打开文件'
        filt = '所有文件(*.*)'
        filename, filtUsed = QFileDialog.getOpenFileName(self, dlgTitle, curPath, filt)
        self.ui.lineEdit_FindFileEmail.setText(filename)

    @pyqtSlot() #一键发送
    def on_btn_FileEmailSend_clicked(self):
        try:
            self.ui.textEdit_emailstatus.clear()
            host_server=self.ui.comboBox_FileEmailsmtpserver.currentText()
            sender_email=self.ui.lineEdit_FileEmailsenderemail.text()
            password=self.ui.lineEdit_FileEmailsenderpassword.text()
            receiver=self.ui.lineEdit_FileEmailreceiveremail.text()
            main_title=self.ui.lineEdit_FileEmailtitle.text()
            main_content=str(self.ui.textEdit_emailContent.toPlainText())
            filename=self.ui.lineEdit_FindFileEmail.text()
            self.__FileEmail(host_server=host_server,sender_email=sender_email,
                            password=password,receiver=receiver,
                            main_title=main_title,main_content=main_content,filename=filename)
            self.ui.textEdit_emailstatus.setPlainText("发送成功！")
        except:
            self.ui.textEdit_emailstatus.clear()
            self.ui.textEdit_emailstatus.setPlainText("发送失败！请检查是否操作正确！")

#  ============窗体测试程序 ================================
#if  __name__ == "__main__":        #用于当前窗体测试
#    app = QApplication(sys.argv)    #创建GUI应用程序
#    form=QmyMainWindow()            #创建窗体
    #icon = QIcon("logo.ico")
    #app.setWindowIcon(icon)
#    form.show()
#    sys.exit(app.exec_())
