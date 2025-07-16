import copy
import re

from settings import (
    TABLE_HEADER,
)


def sort_data_sheets(data_sheets: list) -> list:
    """Складывает строки страниц и сортирует их."""

    yellow_rows = []
    other_rows = []
    double = []
    background_yellow = {'green': 1, 'red': 1}

    for row in data_sheets:
        values = row.get('values')
        fio, phone, *rest = values
        person = (fio, phone)

        if fio and person not in double:
            double.append(person)

            row_yellow = fio['userEnteredFormat'].get('backgroundColor')

            if row_yellow == background_yellow:
                yellow_rows.append(row)
            else:
                other_rows.append(row)

    yellow_rows = sorted(yellow_rows, key=lambda x: x['values'][0]['formattedValue'])
    other_rows = sorted(other_rows, key=lambda x: x['values'][0]['formattedValue'])

    return yellow_rows + other_rows


def format_data_sheets(data_sheets: list, table_end: list) -> list:
    """Корректирует формулы, которые находятся в конце таблицы."""

    ranges = len(data_sheets) + TABLE_HEADER

    for index, row in enumerate(data_sheets, start=1):
        values = row.get('values')
        number = copy.deepcopy(values[2])
        number['formattedValue'] = str(index)
        values.insert(0, number)

    for row_end in table_end:
        for value in row_end['values']:
            user_entered_value = value.get('userEnteredValue')

            if user_entered_value:
                formula = user_entered_value.get('formulaValue')
                match = re.match(
                    r'^=SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)$',
                    formula,
                    re.IGNORECASE,
                )

                if match:
                    column = match.group(1)
                    value['userEnteredValue']['formulaValue'] = f"=SUM({column}4:{column}{ranges})"

                match_for_mini_table = re.match(
                    (
                        r'^=COUNTIFS\(\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+);'
                        r'([^;]+);\$([A-Z]+)\$(\d+):\$([A-Z]+)\$(\d+);"([^"]+)"\)$'
                    ),
                    formula,
                    re.IGNORECASE,
                )

                if match_for_mini_table:
                    column = match_for_mini_table.group(1)
                    next_column = match_for_mini_table.group(6)
                    letter = match_for_mini_table.group(10)
                    value['userEnteredValue']['formulaValue'] = (
                        f'=COUNTIFS(${column}$4:${column}${ranges};1;'
                        f'${next_column}$4:${next_column}${ranges};"{letter}")'
                    )
                    break

    return data_sheets + table_end
