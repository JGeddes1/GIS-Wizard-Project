import sys
from PyQt5 import QtGui
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QAction, QApplication, QDialog, QDialogButtonBox, QHeaderView, QMainWindow, QListWidget, QListWidgetItem, QTableWidget, QTableWidgetItem, QPushButton, QVBoxLayout, QWidget, QFileDialog, QLabel, QMessageBox
from PyQt5.QtGui import QBrush, QClipboard, QColor
import os
import geopandas as gpd
import pandas as pd
import os
import matplotlib.pyplot as plt
from PIL import Image

def printGIS(input_dir,output_dir):
    imageCreation = CreateImages

    imageCreation.generate_preview(input_dir, output_dir)
    

def getFirstSheetName(excel_path):
    xls = pd.ExcelFile(excel_path)
    return xls.sheet_names[0] if xls.sheet_names else None

def createMetadataSheetTable(excel_path, sheet_name):

    
    metadata_df_visual_table = pd.read_excel(excel_path, sheet_name=sheet_name, header=[0, 1,2])
    print("DOING metadata df visual table creation")
    return metadata_df_visual_table


def scan_directory_for_shp(dir_path, excel_path, sheet_name):
    metadata_df = pd.read_excel(excel_path, sheet_name=sheet_name)

    print(metadata_df)
    headerList = ['Table Attribute (each new attribute/notation should be added as a new row)', 'Attributes', 'Table Attribute', 'Attribute', "Attribute "]
    excel_codes = None
    for header in headerList:
        if header in metadata_df:
            print("this header is in " +header)
            excel_codes = metadata_df[header].dropna().unique().tolist()
            break
    if excel_codes is None:
        excel_codes = "Attribute"
        # raise ValueError("None of the specified columns were found in the Excel sheet")   
        print("ARGGGGGGGGGGGGGGGG") 
    # excel_codes = metadata_df['Table Attribute (each new attribute/notation should be added as a new row)'].dropna().unique().tolist()
    print(excel_codes)
    results = {}
    all_matching_headers = []
    for filename in os.listdir(dir_path):
        if filename.endswith(".shp"):
            shp_path = os.path.join(dir_path, filename)
            gdf = gpd.read_file(shp_path)
            shp_headers = list(gdf.columns)
            matching_headers = list([header for header in shp_headers if header.lower() in map(str.lower, excel_codes) and header != "geometry"])

            not_matching_headers = [header for header in shp_headers if header.lower() not in map(str.lower, excel_codes) and header != "geometry"]
            results[filename] = (not_matching_headers, gdf)
            all_matching_headers.extend(matching_headers)
            print("this is all matching_headers  "+str(all_matching_headers))
    return results, all_matching_headers

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('GIS Shapefile Header Analysis Tool')
        self.setGeometry(100, 100, 800, 600)
        self.shapefile_tables = {}
        
        self.setWindowIcon(QtGui.QIcon('c:\\python scripts\\catwizardicon.ico'))
        help_menu = self.menuBar().addMenu("&About")
        about_action = QAction(  "About This App", self)
        about_action.setStatusTip("Find out more about This app and the creator") 
        about_action.triggered.connect( self.about )
        help_menu.addAction(about_action)
        layout = QVBoxLayout()
        widget = QWidget()
        widget.setLayout(layout)
        self.setCentralWidget(widget)
        
        self.listWidget = QListWidget()
        layout.addWidget(self.listWidget)

        # Adding table to display metadata
        self.metadataTable = QTableWidget()
        layout.addWidget(self.metadataTable)
        
        # Adding labels to display selected directory and Excel file
        self.dirLabel = QLabel("Selected Directory: None")
        layout.addWidget(self.dirLabel)
        self.fileLabel = QLabel("Selected Excel File: None")
        layout.addWidget(self.fileLabel)
        
        # Adding buttons
        self.dirButton = QPushButton("Select Directory")
        self.dirButton.clicked.connect(self.askForDirectory)
        layout.addWidget(self.dirButton)
        
        self.fileButton = QPushButton("Select Excel File")
        self.fileButton.clicked.connect(self.askForExcelFile)
        layout.addWidget(self.fileButton)
        
        # Adding button to copy non-matching headers
        self.copyButton = QPushButton("Copy Non-Matching Headers")
        self.copyButton.clicked.connect(self.copyNonMatchingHeaders)
        layout.addWidget(self.copyButton)

         # Adding button to copy non-matching headers
        self.printPreviewsButton = QPushButton("Print preview and thumbs")
        self.printPreviewsButton.clicked.connect(self.gisImagesLoad)
        layout.addWidget(self.printPreviewsButton)
        
        self.listWidget.itemClicked.connect(self.showTable)
    
    def about(self):
        dlg = AboutDialog()
        dlg.exec_()


    def askForDirectory(self):
        dir_path = QFileDialog.getExistingDirectory(self, "Select Directory")
        if dir_path:
            self.dir_path = dir_path
            self.dirLabel.setText(f"Selected Directory: {dir_path}")
            self.tryLoadData()

    def askForExcelFile(self):
        excel_path, _ = QFileDialog.getOpenFileName(self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)")
        if excel_path:
            self.excel_path = excel_path
            self.fileLabel.setText(f"Selected Excel File: {excel_path}")
            self.tryLoadData()

    def tryLoadData(self):
        if hasattr(self, 'dir_path') and hasattr(self, 'excel_path'):
            sheet_name = getFirstSheetName(self.excel_path)
            self.shapefile_tables, self.all_matching_headers = scan_directory_for_shp(self.dir_path, self.excel_path, sheet_name)
            self.metadata_df_visual_table = createMetadataSheetTable(self.excel_path, sheet_name)
            self.updateList()
            self.showMetadata()

    def gisImagesLoad(self):
        if hasattr(self, 'dir_path'):
            output_path = self.dir_path+"\\previews_GIS"
            printGIS(self.dir_path,output_path)
            QMessageBox.information(self, "Copy Non-Matching Headers", "Previews and thumbs available at " + output_path)

    def updateList(self):
        self.listWidget.clear()
        for filename, data in self.shapefile_tables.items():
            not_matching_headers, _ = data
            item_text = f"{filename} - Non-matching headers: {not_matching_headers}" if not_matching_headers else f"{filename} - All headers match"
            item = QListWidgetItem(item_text)
            self.listWidget.addItem(item)
            

    def showMetadata(self):
        if hasattr(self, 'metadata_df_visual_table'):
            self.metadata_df_visual_table.dropna(how='all', inplace=True)
            self.metadataTable.setRowCount(len(self.metadata_df_visual_table))
            self.metadataTable.setColumnCount(len(self.metadata_df_visual_table.columns))
            
            headers = []
            
            for col in self.metadata_df_visual_table.columns:
                if isinstance(col, tuple):
                    # If there are unnamed columns, replace them with an empty string
                    cleaned_col = [elem if 'Unnamed' not in str(elem) else '' for elem in col]
                    headers.append('\n'.join(cleaned_col).strip())
                else:
                    headers.append(col)
            
            self.metadataTable.setHorizontalHeaderLabels(headers)
            
            for index, item in enumerate(headers):
                print(index, item)
                self.metadataTable.resizeColumnToContents(index)
            
            for row_index, row_data in self.metadata_df_visual_table.iterrows():
                            for col_index, value in enumerate(row_data):
                                item = QTableWidgetItem(str(value))
                                lst = [x.lower() for x in self.all_matching_headers]
                                
                                # Highlight matching fields in green
                                if isinstance(value, str) and value.lower() in lst:
                                    item.setBackground(QBrush(QColor(144, 238, 144)))  # Light green background
                                self.metadataTable.setItem(row_index, col_index, item)



                            # Style the headers
        header = self.metadataTable.horizontalHeader()
        header.setStyleSheet("QHeaderView::section { background-color: lightblue; }")
        header.setDefaultAlignment(Qt.AlignCenter)
            # Set alternating row colors
        self.metadataTable.setAlternatingRowColors(True)
        self.metadataTable.setStyleSheet("alternate-background-color: #f0f0f0;")
            
       



    def showTable(self, item):
        filename = item.text().split(' - ')[0]  # Get the filename part
        _, gdf = self.shapefile_tables[filename]
        self.tableWidget = QTableWidget()
        self.tableWidget.setRowCount(len(gdf))
        self.tableWidget.setColumnCount(len(gdf.columns))
        self.tableWidget.setHorizontalHeaderLabels(gdf.columns)
        
        for row_index, row_data in gdf.iterrows():
            for col_index, value in enumerate(row_data):
                self.tableWidget.setItem(row_index, col_index, QTableWidgetItem(str(value)))
        
        self.tableWidget.setWindowTitle(f"Attribute Table - {filename}")
        self.tableWidget.setGeometry(100, 100, 800, 600)  # Adjust size and position 
        self.tableWidget.show()

    def displayTable(self, gdf, filename):
        tableWidget = QTableWidget()
        tableWidget.setRowCount(len(gdf))
        tableWidget.setColumnCount(len(gdf.columns))
        tableWidget.setHorizontalHeaderLabels(gdf.columns)
        
        for row_index, row_data in gdf.iterrows():
            for col_index, value in enumerate(row_data):
                tableWidget.setItem(row_index, col_index, QTableWidgetItem(str(value)))
        
        tableWidget.setWindowTitle(f"Attribute Table - {filename}")
        tableWidget.setGeometry(100, 100, 800, 600)
        tableWidget.show()

    def copyNonMatchingHeaders(self):
        non_matching_headers_list = []
        for filename, data in self.shapefile_tables.items():
            not_matching_headers, _ = data
            if not_matching_headers:
                non_matching_headers_list.append(f"{filename}: {not_matching_headers}")
        
        if non_matching_headers_list:
            non_matching_headers_text = "\n".join(non_matching_headers_list)
            clipboard = QApplication.clipboard()
            clipboard.setText(non_matching_headers_text)
            QMessageBox.information(self, "Copy Non-Matching Headers", "Non-matching headers have been copied to clipboard.")
        else:
            QMessageBox.information(self, "Copy Non-Matching Headers", "There are no non-matching headers to copy.")


class AboutDialog(QDialog):

    def __init__(self, *args, **kwargs):
        super(AboutDialog, self).__init__(*args, **kwargs)

        QBtn = QDialogButtonBox.Ok  # No cancel
        self.buttonBox = QDialogButtonBox(QBtn)
        self.buttonBox.accepted.connect(self.accept)
        self.buttonBox.rejected.connect(self.reject)

        layout = QVBoxLayout()

        title = QLabel("GIS Attribute Wizard Checker")
        font = title.font()
        font.setPointSize(20)
        title.setFont(font)

        layout.addWidget(title)


        layout.addWidget( QLabel("Version 1.1.2") )
        layout.addWidget( QLabel("Copyright 2024 Jamie Geddes.") )

        for i in range(0, layout.count() ):
            layout.itemAt(i).setAlignment( Qt.AlignHCenter )

        layout.addWidget(self.buttonBox)

        self.setLayout(layout)


class CreateImages():

    def __init__(input_dir):

        input = input_dir
        output_dir = (input_dir +"\\previews_GIS")

    def generate_preview(input_dir, output_dir):
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        for file_name in os.listdir(input_dir):
                if file_name.endswith(".shp"):
                    file_path = os.path.join(input_dir, file_name)
                    output_path = os.path.join(output_dir, file_name.replace(".shp", ".jpg"))
                    gdf = gpd.read_file(file_path)
                    fig, ax = plt.subplots(dpi=500)
                    gdf.plot(ax=ax)
                    ax.axis('off')
                    plt.savefig((output_path), bbox_inches='tight', pad_inches=0)
        CreateImages.resize_image(output_dir)


    def resize_image(directory):

        # Get a list of all files in the directory
        files = os.listdir(directory)
        final_image_directory = (directory + "\\thumbnail_and_preview")

        if not os.path.exists(final_image_directory):
            os.makedirs(final_image_directory)

        # Iterate over each file in the directory
        for file_name in files:
            # Check if the file is a JPEG image
            if file_name.endswith('.jpg'):
                # Open the image
                with Image.open(os.path.join(directory, file_name)) as img:
                    # Resize image to max length of 750
                    img_750 = img.copy()
                    original_width, original_height = img.size
                    target_size = [750,125]
                                            # Resize image to max length of 125
                    img_125 = img.copy()

                    for size in target_size:
                        if size == 750:

                            scale_factor = size / max(original_width, original_height)
                            new_width = int(original_width * scale_factor)
                            new_height = int(original_height * scale_factor)


                            img_750 = img.resize((new_width, new_height), Image.LANCZOS)
                            # img_750.thumbnail((new_width, new_height), Image.Resampling.LANCZOS)
                            # Save resized image
                            img_750.save(os.path.join(final_image_directory, file_name),quality=500)
                        if size == 125:
                            
                            scale_factor = size / max(original_width, original_height)
                            new_width = int(original_width * scale_factor)
                            new_height = int(original_height * scale_factor)


                            img_125 = img.resize((new_width, new_height), Image.LANCZOS)
                            # img_125.thumbnail((125, 125), Image.Resampling.LANCZOS)
                            # Save resized image
                            img_125.save(os.path.join(final_image_directory, f"thumb_{file_name[:-4]}.jpg"),quality=500)

        print("Images resized successfully.")
        

        
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())


