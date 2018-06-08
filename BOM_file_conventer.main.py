#/usr/bin/python3
"""Converting BOM to KiCad BOM standard

Program was created to export BOM from EasyEDA to PartsBox project creator via BOM"""

#import built-in modules
import sys
import os
import csv
import zipfile
import math

import xlrd

__author__ = "Kornel Stefańczyk"
__license__ = "CC BY-SA"    #cc-by-sa-4.0
__version__ = "0.2"
__maintainer__ = "Kornel Stefańczyk"
__email__ = "kornelstefanczyk@wp.pl"

class BOMData:
    """Class contain data of one part included in BOM"""

    def __init__(self,
            quantity = None,
            manufacturer_code = None,
            id = None,
            part_name = None,
            package = None,
            circut_index = None,
            lcsc_index = None,
            supplier = None,
            manufacturer = None,
            comment = None):

        self.quantity = quantity,
        #Most important parameter, name in PartsBox
        self.manufacturer_code = manufacturer_code
        self.id = id
        self.part_name = part_name
        self.package = package
        self.circut_index = circut_index
        self.lcsc_index = lcsc_index
        self.supplier = supplier
        self.manufacturer = manufacturer
        self.comment = comment

        self.convert_to_int()

    def convert_to_int(self):
        """Convert id and quantity value to integer"""
        #This function is needed, because this val is returned not like int
        if self.id != 'ID':
            self.id = int(float(self.id))
            if type(self.quantity).__name__ == 'tuple':
                self.quantity = self.quantity[0]
            self.quantity = int(float(self.quantity[0]))

    def print_data(self):
        """Print all information about part to stdout"""
        print("ID: " + str(self.id))
        print("Quantity: " + str(self.quantity))
        print("Part name: " + str(self.part_name))
        print("Manufacturer Code: "+ str(self.manufacturer_code))

        print("Package: " + str(self.package))
        print("Circut Index: " + str(self.circut_index))
        print("LCSC Index: " + str(self.lcsc_index))
        print("Supplier: " + str(self.supplier))
        print("Manufacturer: " + str(self.manufacturer))
        print("Comment: " +  str(self.comment))



class BOMContainer:
    """Class contain data of BOM"""

    def __init__(self):
        self.bom_element_list = []
        self.files_to_remove = []
        self.input_file_patch = None
        self.output_file_patch = None

    def read_user_options(self):
        """Read users option like patchs to input and output files

        Display help """
        first_file_patch = None
        second_file_patch = None
        program_name = sys.argv[0]
        arguments = sys.argv[1:]
        count = len(arguments)
        if count >= 1:
            if sys.argv[1] != '--help':
                if count >= 1:
                    first_file_patch = arguments[0]
                if count >= 2:
                    second_file_patch = arguments[1]
            elif sys.argv[1] == '--help':
                print('Program convert BOM list from EasyEDA to KiCad format')
                print('Usage: ' + program_name + ' [input file patch [output file patch]]')
                print('If you don\'t type input or output file patch ', end='')
                print('you will have to do this later')
                quit()

        if not first_file_patch:
            first_file_patch = input('Set localisation of input file: ')
        if not second_file_patch:
            second_file_patch = input('Set localisation of output file: ')

        if first_file_patch:
            self.input_file_patch = first_file_patch
        if second_file_patch:
            self.output_file_patch = second_file_patch
        else:
            self.output_file_patch = "BOM_file_output.csv"


    def csv_read(self, filename, input_data_format='LCSC', remove=False):
        """Read data from CSV file

        input_data_format:
            EasyEDA - data come from EasyEda BOM(directly by EasyEDA
            LCSC    - data transmitted form EasyEDA via LCSC in xlsx file"""
        if input_data_format is 'EasyEDA':
            with open(filename, newline='', encoding='utf-16') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t')
                for row in reader:
                    if str(row[0]) != 'id':
                        self.bom_element_list.append(BOMData(
                                                id = row[0],
                                                quantity = row[2],
                                                manufacturer_code = row[5],
                                                part_name = row[1],
                                                package = row[3],
                                                circut_index = row[4],
                                                lcsc_index = row[6],
                                                supplier = row[7],
                                                manufacturer = row[8]))

        if input_data_format is 'LCSC':
            with open(filename, newline='', encoding='utf-16') as csvfile:
                reader = csv.reader(csvfile, delimiter='\t')
                for row in reader:
                    if str(row[0]) != 'id':
                        self.bom_element_list.append(BOMData(
                                                id = row[0],
                                                quantity = row[4],
                                                manufacturer_code = row[5],
                                                part_name = row[1],
                                                package = row[3],
                                                circut_index = row[2],
                                                lcsc_index = row[8],
                                                supplier = row[7],
                                                manufacturer = row[6]))

        if remove:
            self.files_to_remove.append(filename)

    def csv_write(self, filename, output_data_format='KiCad',
            description_row=True):
        """Write data into CSV file

        output_data_format:
            KiCad - data to KiCad format(free format in PartsBox)"""
        if output_data_format is 'KiCad':
            with open(filename, 'w', newline='') as csvfile:
                writer = csv.writer(csvfile, delimiter=';',
                                    quoting = csv.QUOTE_ALL)
                if description_row:
                    writer.writerow(['Id', 'Designator', 'Package', 'Quantity',
                        'Designation', 'Supplier and ref'])
                for row in self.bom_element_list:
                    writer.writerow([row.id, row.part_name, row.package,
                        row.quantity, row.manufacturer_code, row.circut_index])


    def unzip_file(self, org_file_path=None, dest_path=None, remove=False):
        """Unzipping zip file, not in use, because EasyEDA change way of
        distributing BOM file."""
        if org_file_path == None:
            if self.input_file_patch.endswith(".zip"):
                org_file_path  = self.input_file_patch
        if org_file_path:
            with zipfile.ZipFile(org_file_path, "r") as zip_ref:
                zip_ref.extractall("tmp_BOM_extract")
            if remove:
                self.files_to_remove.append(org_file_path)

    def csv_from_excel(self, input_file=None, output_file=None, remove=False):
        """Convert xlsx file from LCSC to old BOM csv format used by EasyEDA"""
        if not input_file:
            input_file = self.input_file_patch
        if not output_file:
            output_file = "tmp_BOM_csv_file_converted_from_xlsx.csv"

        wb = xlrd.open_workbook(input_file)
        sh = wb.sheet_by_name('sheet1')
        your_csv_file = open(output_file, 'w', newline='', encoding='utf-16')
        wr = csv.writer(your_csv_file,  delimiter='\t')

        for rownum in range(sh.nrows):
            wr.writerow(sh.row_values(rownum))

        your_csv_file.close()
        if remove:
            self.files_to_remove.append(input_file)
        return output_file

    def print_data(self):
        """Print data to stdout"""
        for i in self.bom_element_list:
            i.print_data()

    def remove_files(self):
        """Remove files mentioned in files_to_remove list"""
        for i in self.files_to_remove:
            os.remove(i)

    def main(self):
        """Execute needed comands to successfully convert BOM file to
        PartsBox foramt"""
        self.read_user_options()
        filename = self.csv_from_excel(remove=False)
        self.csv_read(filename=filename, remove=True)
        self.print_data()
        self.csv_write(filename=self.output_file_patch,description_row=False)
        self.remove_files()


BOM = BOMContainer()
BOM.main()
