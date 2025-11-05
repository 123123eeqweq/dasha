import xlrd
import pandas as pd

def read_excel_file(file_path):
    """Читает Excel файл и выводит его содержимое"""
    try:
        # Открываем файл
        workbook = xlrd.open_workbook(file_path)
        
        print(f"Файл: {file_path}")
        print(f"Количество листов: {workbook.nsheets}")
        print("-" * 50)
        
        # Проходим по всем листам
        for sheet_idx in range(workbook.nsheets):
            sheet = workbook.sheet_by_index(sheet_idx)
            print(f"\nЛист {sheet_idx + 1}: '{sheet.name}'")
            print(f"Строк: {sheet.nrows}, Столбцов: {sheet.ncols}")
            print("-" * 50)
            
            # Выводим первые 20 строк для примера
            max_rows = min(20, sheet.nrows)
            for row_idx in range(max_rows):
                row_values = []
                for col_idx in range(sheet.ncols):
                    cell_value = sheet.cell_value(row_idx, col_idx)
                    row_values.append(str(cell_value))
                print(f"Строка {row_idx + 1}: {' | '.join(row_values)}")
            
            if sheet.nrows > 20:
                print(f"... и еще {sheet.nrows - 20} строк")
        
        # Также можно использовать pandas для более удобной работы
        print("\n" + "=" * 50)
        print("Чтение через pandas:")
        print("=" * 50)
        
        df = pd.read_excel(file_path, engine='xlrd')
        print(df.head(10))  # Первые 10 строк
        print(f"\nРазмер: {df.shape[0]} строк, {df.shape[1]} столбцов")
        
    except Exception as e:
        print(f"Ошибка при чтении файла: {e}")

if __name__ == "__main__":
    # Читаем файл excel.xls
    read_excel_file("excel.xls")

