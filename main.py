from helpers import (
    sort_data_sheets,
    format_data_sheets,
)
from settings import (
    service,
    GENERAL_STREADSHEET_ID,
    LIST_SPREADSHEET_IDS,
    TABLE_HEADER,
)


class GoogleTable(object):
    RANGES_ROWS = "!A4:A"
    END_COLUMN_INDEX = 2
    TABLE_END_COLUMN_INDEX = 7

    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id
        self.spreadsheet = self._get_spreadsheet()
        self.sheets = self._get_sheets()
        self.table_end = None

    def _get_spreadsheet(self):
        return service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            includeGridData=True,
        ).execute()

    def _get_sheets(self):
        return self.spreadsheet.get('sheets')

    def _get_table_end(self, title_sheets: str, len_rows: int):
        """Получает последние строки таблицы с формулами."""

        ranges = TABLE_HEADER + len_rows

        results = service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            ranges=f"{title_sheets}!A{ranges + 1}:K{ranges + self.TABLE_END_COLUMN_INDEX}",
            fields="sheets(data(rowData(values(formattedValue,userEnteredFormat,userEnteredValue(formulaValue)))))",
        ).execute()

        self.table_end = results['sheets'][0]['data'][0].get('rowData')

    def _get_range_for_data(self, title_sheets: str) -> str:
        """Вычисляет диапазон заполненных строк."""

        rows_results = service.spreadsheets().values().batchGet(
            spreadsheetId=self.spreadsheet_id,
            ranges=title_sheets + self.RANGES_ROWS,
            valueRenderOption='FORMATTED_VALUE',
            dateTimeRenderOption='FORMATTED_STRING',
        ).execute()

        len_rows = len(rows_results['valueRanges'][0].get('values', []))

        if len_rows:
            result = f"{title_sheets}!B4:K{str(TABLE_HEADER + len_rows)}"
            self._get_table_end(title_sheets, len_rows)
        else:
            result = f"{title_sheets}!A4:BB"

        return result

    def get_data_with_formatting(self, data_sheet: dict):
        """Получает данные заполненных строк вместе с форматом."""

        title_sheets = str(data_sheet['properties']['title'])
        ranges = self._get_range_for_data(title_sheets)

        results = service.spreadsheets().get(
            spreadsheetId=self.spreadsheet_id,
            ranges=ranges,
            fields="sheets(data(rowData(values(formattedValue,userEnteredFormat))))",
        ).execute()

        return results['sheets'][0]['data'][0]['rowData']

    def clear_sheet(self, range: str):
        service.spreadsheets().values().clear(
            spreadsheetId=self.spreadsheet_id,
            range=range,
            body={}
        ).execute()

    def rewrite_table(self, sheet_info: dict, sheet_data_with_format: list):
        requests = []
        title_sheets = str(sheet_info['properties']['title'])
        ranges = self._get_range_for_data(title_sheets)
        self.clear_sheet(ranges)

        for i, row_data in enumerate(sheet_data_with_format):
            for j, cell_data in enumerate(row_data.get('values', [])):
                cell_format = cell_data.get('userEnteredFormat', {})
                cell_value = cell_data.get('userEnteredValue', {})
                user_entered_value = {}

                if 'formulaValue' in cell_value:
                    user_entered_value['formulaValue'] = cell_value['formulaValue']
                else:
                    if cell_data.get('formattedValue', '').isdigit():
                        user_entered_value['numberValue'] = cell_data.get('formattedValue', '')
                    else:
                        user_entered_value['stringValue'] = cell_data.get('formattedValue', '')

                requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_info['properties']['sheetId'],
                            "startRowIndex": TABLE_HEADER + i,
                            "startColumnIndex": j,
                            "endColumnIndex": self.END_COLUMN_INDEX + j,
                        },
                        "rows": [{
                            "values": [{
                                "userEnteredValue": user_entered_value,
                                "userEnteredFormat": cell_format,
                            }]
                        }],
                        "fields": "userEnteredValue,userEnteredFormat",
                    }
                })

        if requests:
            service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheet_id,
                body={"requests": requests}
            ).execute()


def merge_data_table(list_spreadsheet_ids: list):
    general_table = GoogleTable(GENERAL_STREADSHEET_ID)

    for sheet_num in range(0, len(general_table.sheets)):
        data_sheets = []
        table_list = []

        for spreadsheet_id in list_spreadsheet_ids:
            table = GoogleTable(spreadsheet_id)
            data_sheets = data_sheets + table.get_data_with_formatting(table.sheets[sheet_num])
            table_list.append(table)

        data_sheets = sort_data_sheets(data_sheets)
        data_sheets = format_data_sheets(data_sheets, table_list[0].table_end)
        general_table.rewrite_table(general_table.sheets[sheet_num], data_sheets)

    return print("Слияние таблиц выполнено.")


merge_data_table(LIST_SPREADSHEET_IDS)
