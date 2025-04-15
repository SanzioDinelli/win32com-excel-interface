import win32com.client
from time import sleep

class ExcelManager:
    def __init__(self, file_path=None, visible=True):
        """Inicializa o Excel e abre um arquivo opcionalmente."""
        try:
            # Tenta conectar-se a uma instância já aberta do Excel
            self.excel = win32com.client.GetActiveObject("Excel.Application")
            print("Conectado ao Excel existente.")
            self.workbook = self.excel.ActiveWorkbook
            self.workbook = self.excel.Workbooks.Add()
        except Exception:
            # Se não houver uma instância, cria uma nova
            self.excel = win32com.client.Dispatch("Excel.Application")
            print("Nova instância do Excel iniciada.")

            if file_path:
                self.workbook = self.excel.Workbooks.Open(file_path)
            else:
                self.workbook = self.excel.Workbooks.Add()
        
        self.excel.Visible = visible  # Exibir ou não o Excel

    def open_workbook(self, file_path):
        """Abre um arquivo Excel."""
        self.workbook = self.excel.Workbooks.Open(file_path)

    def get_sheet(self, sheet_index=1):
        """Retorna uma planilha específica (padrão: primeira)."""
        return self.workbook.Sheets(sheet_index)

    def get_value(self, sheet_index, row, col):
        """Obtém o valor de uma célula específica."""
        sheet = self.get_sheet(sheet_index)
        return sheet.Cells(row, col).Value

    def set_value(self, sheet_index, row, col, value):
        """Define o valor de uma célula específica."""
        sheet = self.get_sheet(sheet_index)
        sheet.Cells(row, col).Value = value

    def get_row_count(self, sheet_index):
        """Obtém a quantidade de linhas usadas na planilha."""
        sheet = self.get_sheet(sheet_index)
        return sheet.UsedRange.Rows.Count

    def get_column_count(self, sheet_index):
        """Obtém a quantidade de colunas usadas na planilha."""
        sheet = self.get_sheet(sheet_index)
        return sheet.UsedRange.Columns.Count

    def get_range_values(self, sheet_index, start_cell, end_cell):
        """Obtém os valores de um intervalo de células."""
        sheet = self.get_sheet(sheet_index)
        return sheet.Range(start_cell, end_cell).Value

    def parse_list_sheetrange(self, list_to_parse):
        """Converte uma lista de itens unitários para uma lista de tuplas para inserção em colunas"""
        sheetrange = []
        for i in list_to_parse:
            sheetrange.append((i,))
        return sheetrange
    
    def set_range_values(self, sheet_index: int | str, start_cell: tuple, end_cell: tuple, values: list):
        """Define valores em um intervalo de células."""
        sheet = self.get_sheet(sheet_index)
        start_cell = sheet.Cells(start_cell[0],start_cell[1])
        end_cell = sheet.Cells(end_cell[0],end_cell[1])
        sheet.Range(start_cell, end_cell).Value = values

    def insert_row(self, sheet_index, row):
        """Insere uma nova linha na posição especificada."""
        sheet = self.get_sheet(sheet_index)
        sheet.Rows(row).Insert()

    def insert_column(self, sheet_index, col):
        """Insere uma nova coluna na posição especificada."""
        sheet = self.get_sheet(sheet_index)
        sheet.Columns(col).Insert()

    def delete_row(self, sheet_index, row):
        """Remove uma linha específica."""
        sheet = self.get_sheet(sheet_index)
        sheet.Rows(row).Delete()

    def delete_column(self, sheet_index, col):
        """Remove uma coluna específica."""
        sheet = self.get_sheet(sheet_index)
        sheet.Columns(col).Delete()

    def insert_image(self, sheet_index, image_path, left, top, width, height):
        """Insere uma imagem na planilha."""
        sheet = self.get_sheet(sheet_index)
        sheet.Shapes.AddPicture(image_path, 1, 1, left, top, width, height)
    
    def insert_image_over_range(self, sheet_index, image_path, cell): # Selection.InsertPictureInCell
        """Insere uma imagem na planilha."""
        sheet = self.get_sheet(sheet_index)

        img = sheet.Pictures().Insert(image_path)
        cell_fixed = sheet.Range(cell)  # Define a célula onde deseja inserir
        img.Top = cell_fixed.Top
        img.Left = cell_fixed.Left
        img.Width = cell_fixed.Width
        img.Height = cell_fixed.Height
        img.Placement = 1

    def insert_image_on_cell(self, image_path, cell, macro_name):
        """Insere uma imagem na célula."""
        # vb_project = self.workbook.VBProject

        # # Deletar todos os módulos onde as macros são armazenadas
        # for i in reversed(range(vb_project.VBComponents.Count)):  # Percorre de trás para frente para evitar erros
        #     vb_component = vb_project.VBComponents(i + 1)  # +1 porque a contagem do VBA começa em 1
        #     if vb_component.Type == 1:  # 1 = Módulo padrão
        #         vb_project.VBComponents.Remove(vb_component)



        macro_name = macro_name.lower()

        with open("macro.txt","w") as file:
            file.write(f"""Sub {macro_name}()
'
' {macro_name} Macro
'

'
    Range("{cell}").Select
    Selection.InsertPictureInCell ( _
        "{image_path}" _
        )
End Sub""")
            file.close()
        
        with open("macro.txt","r") as file:
            vba_code = file.read()
            file.close()
        
        vb_module = self.workbook.VBProject.VBComponents.Add(1)  # 1 = Módulo
        vb_module.CodeModule.AddFromString(vba_code)
        sleep(2)
        
        while True:
            try:
                self.excel.Application.Run(macro_name)
                break
            except Exception as error:
                print(error)
                sleep(1)

    def create_pivot_table(self, sheet_index, source_range, table_destination):
        """Cria uma Tabela Dinâmica (Pivot Table)."""
        sheet = self.get_sheet(sheet_index)
        pivot_cache = self.workbook.PivotCaches().Create(1, sheet.Range(source_range))
        pivot_table = pivot_cache.CreatePivotTable(sheet.Range(table_destination), "PivotTable1")
        return pivot_table

    def auto_fit_columns(self, sheet_index):
        """Ajusta automaticamente a largura das colunas."""
        sheet = self.get_sheet(sheet_index)
        sheet.Cells.EntireColumn.AutoFit()

    def set_formula(self, sheet_index, row, col, formula):
        """Define uma fórmula em uma célula."""
        sheet = self.get_sheet(sheet_index)
        sheet.Cells(row, col).Formula = formula

    def set_cell_color(self, sheet_index, row, col, color):
        """Altera a cor de fundo de uma célula. color 0 ~ 999"""
        sheet = self.get_sheet(sheet_index)
        sheet.Cells(row, col).Interior.Color = color

    def save(self):
        """Salva o arquivo atual."""
        self.workbook.Save()

    def save_as(self, new_path):
        """Salva a planilha em um novo caminho."""
        self.workbook.SaveAs(new_path)

    def close(self, save_changes=False):
        """Fecha a planilha e o Excel."""
        if self.workbook:
            self.workbook.Close(SaveChanges=save_changes)
        self.excel.Quit()


if __name__ == "__main__":
    # Exemplo de uso
    excel_manager = ExcelManager(visible=True)
    excel_manager.set_value(1, 1, 1, "Olá, Excel!")
    print(excel_manager.get_value(1, 1, 1))
    excel_manager.save_as(r"C:\Users\marid\OneDrive\Área de Trabalho\Codes_v1.1\win32com-excel-interface\exemplo.xlsx")
    excel_manager.close(save_changes=True)