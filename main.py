import xml.etree.ElementTree as ET
import pandas as pd
from openpyxl import Workbook
import argparse

# Функция для обработки XML и создания Excel-файла
def process_xml_to_excel(input_file):
    # Парсим XML-файл
    tree = ET.parse(input_file)
    root = tree.getroot()

    # Создаем пустой список для хранения данных
    data = []

    # Создаем словарь для хранения сопоставления network_id с интерфейсами и их label
    network_id_mapping = {}

    # Итерируемся по узлам (nodes) в XML-файле
    for node in root.findall('.//node'):
        node_name = node.get('name')
        interfaces = node.findall('./interface')

        for interface in interfaces:
            interface_name = interface.get('name')
            network_id = int(interface.get('network_id'))
            label = interface.get('label')  # Добавляем параметр label
            other_label = interface.get('other_label')  # Добавляем параметр other_label

            # Создаем ключ, объединяя network_id и первые два сегмента label (если они есть)
            key = f"{network_id}_{label.split('.')[0] if '.' in label else label}"

            # Добавляем сопоставление network_id с интерфейсами и их label (в виде списка) в словарь
            if key in network_id_mapping:
                network_id_mapping[key].append((node_name, interface_name, label, other_label))
            else:
                network_id_mapping[key] = [(node_name, interface_name, label, other_label)]

    # Итерируемся по узлам (nodes) в XML-файле еще раз для создания списка данных
    for node in root.findall('.//node'):
        node_name = node.get('name')
        interfaces = node.findall('./interface')

        for interface in interfaces:
            interface_name = interface.get('name')
            network_id = int(interface.get('network_id'))
            label = interface.get('label')

            # Создаем ключ, объединяя network_id и первые два сегмента label (если они есть)
            key = f"{network_id}_{label.split('.')[0] if '.' in label else label}"

            # Получаем все соответствующие пары node name - interface name - label - other_label для данного ключа
            other_interfaces = network_id_mapping.get(key, [])

            # Создаем новую запись для каждой пары, исключая те, где node name и other node name совпадают
            for other_node_name, other_interface_name, other_label, _ in other_interfaces:
                if node_name != other_node_name:
                    data.append([label, interface_name, node_name, other_node_name, other_interface_name, other_label])

    # Создаем DataFrame из списка данных
    df = pd.DataFrame(data, columns=['маркировка', 'порт', 'оборудование', 'оборудование', 'порт', 'маркировка'])

    # Создаем Excel-файл и получаем объект рабочей книги
    wb = Workbook()
    ws = wb.active

    # Добавляем названия столбцов
    column_names = df.columns.tolist()
    ws.append(column_names)

    # Записываем данные из DataFrame в рабочий лист Excel
    for _, row in df.iterrows():
        ws.append(row.tolist())

    # Настраиваем ширину столбцов на основе длины содержимого
    for column_cells in ws.columns:
        max_length = 0
        column = column_cells[0].column_letter
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width

    # Сохраняем Excel-файл с именем, совпадающим с именем входного файла
    output_file = input_file.split('.')[0] + '.xlsx'
    wb.save(output_file)
    print(f'Данные сохранены в файле {output_file}')

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description='Преобразование XML в Excel.')
    parser.add_argument('input_file', help='Имя входного XML-файла')
    args = parser.parse_args()
    input_file = args.input_file
    process_xml_to_excel(input_file)
