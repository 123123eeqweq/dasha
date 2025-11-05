# -*- coding: utf-8 -*-
import sys
import pandas as pd
import xlrd

# Устанавливаем UTF-8 для вывода в Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def merge_articles_to_excel():
    """Добавляет украинские названия из txt в Excel по совпадению артикулов"""
    
    # 1. Читаем txt файл с артикулами и названиями
    print("Читаю articles_output.txt...")
    articles_dict = {}
    
    try:
        with open('articles_output.txt', 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and '\t' in line:
                    parts = line.split('\t', 1)
                    if len(parts) == 2:
                        article = parts[0].strip()
                        ukrainian_name = parts[1].strip()
                        articles_dict[article] = ukrainian_name
        
        print(f"Загружено {len(articles_dict)} артикулов из txt файла")
    except Exception as e:
        print(f"Ошибка при чтении txt файла: {e}")
        return
    
    # 2. Читаем Excel файл
    print("\nЧитаю excel.xls...")
    try:
        # Читаем через pandas с xlrd engine для старых .xls файлов
        df = pd.read_excel('excel.xls', engine='xlrd')
        
        print(f"Размер Excel: {df.shape[0]} строк, {df.shape[1]} столбцов")
        print(f"Столбцы: {list(df.columns)}")
        
        # Ищем столбец с артикулами (stok kodu)
        article_column = None
        for col in df.columns:
            col_lower = str(col).lower()
            if 'stok' in col_lower and 'kodu' in col_lower:
                article_column = col
                break
        
        if article_column is None:
            print("\nНе найден столбец 'stok kodu'. Доступные столбцы:")
            for i, col in enumerate(df.columns):
                print(f"  {i}: {col}")
            return
        
        print(f"\nНайден столбец с артикулами: '{article_column}'")
        
        # 3. Добавляем новую колонку с украинскими названиями
        df['Українська назва'] = ''
        
        # 4. Ищем совпадения и заполняем колонку
        matched_count = 0
        for idx, row in df.iterrows():
            article = str(row[article_column]).strip()
            
            # Пробуем найти точное совпадение
            if article in articles_dict:
                df.at[idx, 'Українська назва'] = articles_dict[article]
                matched_count += 1
            else:
                # Пробуем найти без пробелов и лишних символов
                article_clean = article.replace(' ', '').replace('-', '').replace('.', '')
                for art_key, name in articles_dict.items():
                    art_key_clean = art_key.replace(' ', '').replace('-', '').replace('.', '')
                    if article_clean == art_key_clean:
                        df.at[idx, 'Українська назва'] = name
                        matched_count += 1
                        break
        
        print(f"\nНайдено совпадений: {matched_count} из {len(df)} строк")
        
        # 5. Сохраняем результат
        output_file = 'excel_with_ukrainian_names.xls'
        
        # Для сохранения в .xls используем xlwt
        try:
            df.to_excel(output_file, index=False, engine='openpyxl')
            print(f"\nФайл сохранен как: {output_file}")
        except:
            # Если openpyxl не работает, пробуем xlwt
            try:
                import xlwt
                # Создаем новую книгу
                wb = xlwt.Workbook()
                ws = wb.add_sheet('Sheet1')
                
                # Записываем заголовки
                for col_idx, col_name in enumerate(df.columns):
                    ws.write(0, col_idx, str(col_name))
                
                # Записываем данные
                for row_idx, (_, row) in enumerate(df.iterrows(), 1):
                    for col_idx, col_name in enumerate(df.columns):
                        value = row[col_name]
                        if pd.notna(value):
                            ws.write(row_idx, col_idx, str(value))
                        else:
                            ws.write(row_idx, col_idx, '')
                
                wb.save(output_file)
                print(f"\nФайл сохранен как: {output_file}")
            except Exception as e:
                print(f"\nОшибка при сохранении: {e}")
                print("Пробую сохранить в .xlsx...")
                output_file = 'excel_with_ukrainian_names.xlsx'
                df.to_excel(output_file, index=False, engine='openpyxl')
                print(f"Файл сохранен как: {output_file}")
        
        # Показываем примеры
        print("\nПримеры совпадений:")
        print("=" * 80)
        sample = df[df['Українська назва'] != ''].head(10)
        for idx, row in sample.iterrows():
            print(f"Артикул: {row[article_column]} -> {row['Українська назва']}")
        
    except Exception as e:
        print(f"Ошибка при работе с Excel: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    merge_articles_to_excel()

