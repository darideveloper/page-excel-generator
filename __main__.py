import os
from dotenv import load_dotenv
import openpyxl


# Env variables
load_dotenv()
TEMPLATE = os.getenv("TEMPLATE")
DOMAIN = os.getenv("DOMAIN")
IMAGES_FOLDER = os.getenv("IMAGES_FOLDER")
EXCEL_SHEET = os.getenv("EXCEL_SHEET")


class PageGenerator():
    def __init__(self, template_name: str):
        """ Initialize PageGenerator object
        
        Args:
            template_name (str): Name of the template to be used
        """
        
        print("Initializing PageGenerator...")
        print(f"Template: {template_name}")
        
        self.template_name = template_name
        
        # Paths
        self.current_folder = os.path.dirname(os.path.abspath(__file__))
        self.templates_folder = os.path.join(self.current_folder, "templates")
        self.excels_folder = os.path.join(self.current_folder, "excels")
        
        # Data variables
        self.excel_data = []
        self.excel_header = []
        
        # Validation data
        self.columns = {
            "redirect": {
                "row": 1,
                "names": [
                    "url",
                    "title",
                    "description",
                    "image url",
                    "site name"
                ]
            }
        }
        self.template_columns = self.columns[self.template_name]
        self.template_columns_row = self.template_columns["row"]
        self.template_columns_names = self.template_columns["names"]
        
        # Load excel data and validate columns
        self.__load_excel_data__()
        self.__save_header__()
        self.__validate_excel_columns__()
        
        # Process data
        self.__replace_images_paths__()
        
    def __load_excel_data__(self):
        """ Read excel data and save in instance
        """
        
        print("Loading excel data...")
        
        # Validate excel path
        excel_path = os.path.join(self.excels_folder, f"{self.template_name}.xlsx")
        if not os.path.exists(excel_path):
            raise FileNotFoundError(f"Excel file not found: {excel_path}")
        
        # Read excel
        wb = openpyxl.load_workbook(excel_path)
        current_sheet = wb[EXCEL_SHEET]
        
        rows = current_sheet.max_row
        columns = current_sheet.max_column

        data = []
        for row in range(1, rows + 1):

            row_data = []
            for column in range(1, columns + 1):
                cell_data = current_sheet.cell(row, column).value
                row_data.append(cell_data)

            data.append(row_data)

        self.excel_data = data
        
    def __save_header__(self):
        """ Save in instance the excel header """
        
        template_row = self.template_columns_row
        excel_header = self.excel_data[template_row - 1]
        self.excel_header = excel_header
        
    def __validate_excel_columns__(self):
        """ Check specific columns and column's order in excel """
        
        print("Validating excel columns...")
        
        # Get current columns
        template_columns_names = self.template_columns_names
        
        # Validete excel header
        for column_name in template_columns_names:
            if column_name not in self.excel_header:
                error = f"Column '{column_name}' not found in excel"
                error += f"\nExcel columns: {self.excel_header}"
                raise ValueError(error)
            
    def __replace_images_paths__(self):
        """ Replace images paths using relative paths and images folder """
        
        # Identify images columns
        images_columns_indexes = []
        for column_name in self.excel_header:
            if "image" in column_name:
                column_index = self.excel_header.index(column_name)
                images_columns_indexes.append(column_index)
        
        # Replace images paths
        for row in self.excel_data[self.template_columns_row:]:
            for column_index in images_columns_indexes:
                image_file = row[column_index]
                new_image_path = f"../{IMAGES_FOLDER}/{image_file}"
                row[column_index] = new_image_path

    def generate_pages(self):
        """ Generate pages using template with and excel data """
        
        template_path = os.path.join(self.templates_folder, f"{self.template_name}.html")
        template_content = open(template_path, "r").read()
        
        # generate each html file with excel data
        for row in self.excel_data[self.template_columns_row:]:
            
            content = template_content
            
            # Replace each cell in template
            for cell in row:
                cell_index = row.index(cell)
                current_column_name = self.excel_header[cell_index]
                
                content = content.replace(f"[{current_column_name}]", cell)
                
            with open("temp.html", "w") as f:
                f.write(content)
    
            print()
            
    
if __name__ == "__main__":
    pg = PageGenerator(TEMPLATE)
    pg.generate_pages()