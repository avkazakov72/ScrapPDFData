import pdfplumber
import pandas as pd
import re
from datetime import datetime

def parse_bank_statement(pdf_path):
    # Регулярные выражения
    # !!!
    operation_header = re.compile(
        r'^(\d{2}\.\d{2}\.\d{4})\s+\d{2}:\d{2}\s+(\d{6})\s+(.*?)\s+([+-]?\s*[\d\s]+,\d{2})\s+([\d\s]+,\d{2})$'
    )
    description_line = re.compile(r'^\d{2}\.\d{2}\.\d{4}\s+(.*?)(?:\. Операция по карте|$)')
    page_footer = re.compile(r'Продолжение на следующей странице|Для проверки подлинности документа')
    
    operations = []
    current_op = {}
    in_operations_section = False
    
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            
            for line in lines:
                # Начало раздела операций
                if "Расшифровка операций" in line:
                    in_operations_section = True
                    continue
                    
                if not in_operations_section:
                    continue
                
                # Конец данных на странице
                if page_footer.search(line):
                    break
                
                # Обработка строки с операцией (красный, зеленый, синий, черный)
                header_match = operation_header.match(line)
                if header_match:
                    if current_op:  # Сохраняем предыдущую операцию
                        operations.append(current_op)
                    
                    date_str, auth_code, category, amount, balance = header_match.groups()
                    
                    # Форматируем данные
                    current_op = {
                        'Дата операции': datetime.strptime(date_str, '%d.%m.%Y').strftime('%d.%m.%Y'),
                        'Категория': category.strip(),
                        'Описание операции': '',
                        'Сумма операции': (
                            amount.replace(' ', '').replace(',', '.') 
                            if amount.startswith('+') 
                            else f"-{amount.lstrip('-').replace(' ', '').replace(',', '.')}"
                        ),
                        'Сальдо': balance.replace(' ', '').replace(',', '.')
                    }
                    continue
                
                # Обработка описания операции (желтый)
                if current_op and not operation_header.match(line):
                    desc_match = description_line.match(line)
                    if desc_match:
                        desc_text = desc_match.group(1).strip()
                        if desc_text:
                            if current_op['Описание операции']:
                                current_op['Описание операции'] += ' ' + desc_text
                            else:
                                current_op['Описание операции'] = desc_text
    
    # Добавляем последнюю операцию
    if current_op:
        operations.append(current_op)
    
    # Создаем DataFrame
    df = pd.DataFrame(operations)
    
    if not df.empty:
        # Удаляем строки без ключевых данных
        df = df[
            (df['Дата операции'].notna()) & 
            (df['Сумма операции'] != '') & 
            (df['Сальдо'] != '')
        ]
        
        # Удаляем дубликаты
        df = df.drop_duplicates()
    
    return df

# Использование
pdf_path = "~/PycharmProjects/ScrapDataProject/Data/pdf_data.pdf"
df = parse_bank_statement(pdf_path)

if not df.empty:
    # Сохраняем в Excel
    output_path = "statement_final.xlsx"
    df.to_excel(output_path, index=False)
    
    print(f"Данные успешно сохранены в {output_path}")
    print("\nПример извлеченных данных:")
    print(df.head().to_string(index=False))
else:
    print("Не удалось извлечь данные операций")