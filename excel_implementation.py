from excel_abstraction import EmployeeAbstraction
from excel_classes import Address, Employee
from excelfile import get_wb_and_sheet,file_path
from openpyxl.styles import Alignment





class empimplementation(EmployeeAbstraction):
    
    def writeheader(self):
        workbook,sheet = get_wb_and_sheet()
        header_line = "id name age salary address".split()
        for i in range(1,len(header_line)+1):
            sheet.cell(row=1,column=i).value=header_line[i-1]
        sheet.merge_cells('E1:G1')
        
        address_header_line = "pincode city state".split()
        for i in range(5,8):
            sheet.cell(row=2,column=i).value=address_header_line[i-5]
        
        for i in range(1,sheet.max_row+1):
            for j in range(1,sheet.max_column+1):
                cell = sheet.cell(row=i,column=j)
                cell.alignment = Alignment(horizontal="center",vertical="center")
        workbook.save(file_path)

        
    def writedata(self):
        pass

    
    def readdata(self):
        pass



if __name__=="__main__":
    empl = empimplementation()
    # empl.writeheader()















