# -*- coding: utf-8 -*-
import sys
import pandas as pd

# Устанавливаем UTF-8 для вывода в Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8')

def check_excel_result():
    """Проверяет результат Excel файла и показывает артикулы без совпадений"""
    
    try:
        # Пробуем прочитать файл - сначала как .xls, потом как .xlsx
        df = None
        try:
            df = pd.read_excel('excel_with_ukrainian_names.xls', engine='xlrd')
        except:
            try:
                df = pd.read_excel('excel_with_ukrainian_names.xls', engine='openpyxl')
            except:
                df = pd.read_excel('excel_with_ukrainian_names.xlsx', engine='openpyxl')
        
        if df is None:
            print("Не удалось прочитать файл!")
            return
        
        print(f"Всего строк в Excel: {len(df)}")
        print(f"Столбцы: {list(df.columns)}")
        print("\n" + "=" * 80)
        
        # Находим столбец с артикулами
        article_column = None
        for col in df.columns:
            col_lower = str(col).lower()
            if 'stok' in col_lower and 'kodu' in col_lower:
                article_column = col
                break
        
        if article_column is None:
            print("Не найден столбец с артикулами!")
            return
        
        # Находим столбец с украинскими названиями
        ukr_column = None
        for col in df.columns:
            if 'українська' in str(col).lower() or 'украинск' in str(col).lower():
                ukr_column = col
                break
        
        if ukr_column is None:
            print("Не найден столбец с украинскими названиями!")
            return
        
        # Подсчитываем статистику
        with_name = df[df[ukr_column].notna() & (df[ukr_column] != '')]
        without_name = df[df[ukr_column].isna() | (df[ukr_column] == '')]
        
        print(f"\nСтатистика:")
        print(f"  - С украинским названием: {len(with_name)} строк")
        print(f"  - БЕЗ украинского названия: {len(without_name)} строк")
        
        # Показываем артикулы без совпадений
        if len(without_name) > 0:
            print("\n" + "=" * 80)
            print("АРТИКУЛЫ БЕЗ УКРАИНСКОГО НАЗВАНИЯ (не нашлось совпадение):")
            print("=" * 80)
            
            for idx, row in without_name.iterrows():
                article = str(row[article_column]).strip()
                product_name = str(row.get('ÜRÜN ADI', '')).strip()
                print(f"Артикул: {article} | Название: {product_name}")
        
        # Показываем несколько примеров с совпадениями
        if len(with_name) > 0:
            print("\n" + "=" * 80)
            print("Примеры с совпадениями:")
            print("=" * 80)
            for idx, row in with_name.head(5).iterrows():
                article = str(row[article_column]).strip()
                ukr_name = str(row[ukr_column]).strip()
                print(f"Артикул: {article}")
                print(f"  Украинское название: {ukr_name}")
                print()
        
    except Exception as e:
        print(f"Ошибка: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_excel_result()

