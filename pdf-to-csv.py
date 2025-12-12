#!/usr/bin/python3
import os
import csv
import sys
import pdfplumber
from PyQt5.QtCore import *
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *

class Pdftocsv(QWidget):
	def __init__(self):
		super().__init__()
		self.setWindowTitle("convert tables to csv")
		self.setGeometry(500,0,400,400)
		self.layout = QVBoxLayout()

		# making the two main widgets
		self.button = QPushButton("select pdf")
		self.button.clicked.connect(self.open_file)
		self.text = QTextEdit("Convert your pdf tables into csv file")
		self.text.setReadOnly(True)
		self.text.setFont(QFont("Times",16))

		# progress bar for nowing how far operations are left
		self.progress = QProgressBar()
		self.progress.setValue(0)
		self.progress.setVisible(False)

		# add the widgets to the main window
		self.widgets = [self.text,self.button,self.progress]
		for widget in self.widgets:
			self.layout.addWidget(widget)

		# set the main window layout 
		self.setLayout(self.layout)

	def open_file(self):
		self.file_path, _ = QFileDialog.getOpenFileName(self,"select pdf","","Pdf Files (*.pdf)")
		if self.file_path:
			self.convert_pdf_to_csv(self.file_path)
		else:
			self.text.setText("please select a file")
	def convert_pdf_to_csv(self,file_path):
		self.pdf_path = file_path
		self.csv_path = self.file_path.replace(".pdf",".csv")

		# file overwrite protection
		if os.path.exists(self.csv_path):
			repy = QMessageBox.question(self,
												"File Exists!",
												f"{self.csv_path} already exists Do you want to overwrite it?",
												QMessageBox.Yes | QMessageBox.No
												)
			if repy == QMessageBox.No:
				# allow the user to set a save name
				self.csv_path, _ = QFileDialog.getSaveFileName(
																				self,
																				"Save CSV As",
																				self.csv_path,
																				"CSV Files (*.csv)"
				)

				# if user cancelled then
				if not self.csv_path:
					self.text.setText("no operation done cancelled by the user")
					return
		try:
			with pdfplumber.open(self.pdf_path) as pdf_file:
				# set up progress bar
				self.total_pages = len(pdf_file.pages)
				self.progress.setValue(0)
				self.progress.setRange(0,self.total_pages)
				self.progress.setVisible(True)
				with open(f"{self.csv_path}","w", newline="", encoding="utf-8") as csv_file:
					writer = csv.writer(csv_file)

					# begin blumbing for information
					for page_index,page  in enumerate(pdf_file.pages, start=1):
						tables = page.extract_tables()
						if tables:
							for table_index, table in enumerate(tables, start=1):
								for row in table:
									writer.writerow(row)

								# update the progress bar everytime
								self.progress.setValue(page_index)
								QApplication.processEvents() # keeps gui from freezing
								# write a table separator
								#writer.writerow([f"-----------------------End of table {table_index}--Page {page_index}--------------"])
		except Exception as e:
			self.text.setText(f"erro occured {e}")
		
		# if all went well the coversion ends here
		self.text.setText(f"finished saved the csv file to {self.csv_path}")
		# hide the progress bar
		self.progress.setVisible(False)

def main():
	app = QApplication(sys.argv)
	window = Pdftocsv()
	window.show()
	sys.exit(app.exec_())
if __name__ == "__main__":
	main()
