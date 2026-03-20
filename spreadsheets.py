from modules import *

class Spreadsheet():
    
    def __init__(self, service_email: str, credentials_filename: str, spreadsheet_id: str = ""):    
        """
        Управление таблицей
        Если идентификатор не указан, то будет создана новая таблица.
        """

        self.credentials: ServiceAccountCredentials | None = None
        self.spreadsheetId: str = spreadsheet_id
        self.credentials_filename: str = credentials_filename

    def start(self):
        self.credentials = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_filename, ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"])
        
        # Авторизуемся в системе
        self.httpAuth = self.credentials.authorize(httplib2.Http())

        # Выбираем работу с таблицами и 4 версию API 
        self.service = discovery.build("sheets", "v4", http = self.httpAuth)

        # Проверяем создана ли таблица
        if self.spreadsheetId:
            self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()

    def create_spreadsheet(self, title: str, sheet: str):
        """
        Создаёт гугл таблицу и возвращает её идентификатор.
        """
        spreadsheet = self.service.spreadsheets().create(body = {
        "properties": {"title": title, "locale": "ru_RU"},
        "sheets": [{"properties": {"sheetType": "GRID",
                    "sheetId": 0,
                    "title": sheet,
                    # TODO: Выбор размера таблицы
                    # FIXME: вывод ошибки о переполнении таблицы
                    "gridProperties": {"rowCount": ROW_COUNT, "columnCount": COLUMN_COUNT}}}]
        }).execute()

        self.spreadsheetId = spreadsheet["spreadsheetId"]
    
    def share_spreadsheet(self, gmail: str):
        """
        Выдаёт доступ к таблице пользователю с указанной гугл почтой.
        """
        driveService = discovery.build("drive", "v3", http = self.httpAuth) # Выбираем работу с Google Drive и 3 версию API
        return driveService.permissions().create(
            fileId = self.spreadsheetId,
            body = {"type": "user", "role": "writer", "emailAddress": gmail},  # Открываем доступ на редактирование
            fields = "id"
        ).execute()

    def get(self, address: str):
        """
        Возвращает значение ячейки.

        :param address: Например "B1".
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId,
            range=address
        ).execute()
    
        values = result.get("values", [])
        if values:
            return values[0][0]  # Возвращаем значение первой ячейки в диапазоне
        return None

    def get_row(self, row_number: int) -> list:
        """
        Возвращает все значения из указанной строки.

        :param row_number: Номер строки (начинается с 1).
        :return: Список значений строки.
        """
        range_name = f"{row_number}:{row_number}"  # Например, "1:1" для первой строки
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId,
            range=range_name
        ).execute()

        values = result.get("values", [])
        if values:
            return values[0]  # Возвращаем первую строку (все её значения)
        return []

    def get_column(self, column_letter: str) -> list:
        """
        Возвращает все значения из указанного столбца.

        :param column_letter: Буква столбца (например, "A" или "B").
        :return: Список значений столбца.
        """
        range_name = f"{column_letter}:{column_letter}"  # Например, "A:A" для столбца A
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId,
            range=range_name
        ).execute()

        values = result.get("values", [])
        # Возвращаем все значения столбца, пропуская пустые строки
        return [row[0] if row else "" for row in values]

    def get_column_index(self, sheet_title:str, column_index: int) -> list:
        """
        Возвращает все значения из указанного столбца.

        :param column_index: начиная с 1.
        :return: Список значений столбца.
        """
        sheet: list[list[str]] = self.get_sheet(sheet_title)
        result: list[str] = []

        for row in sheet:
            result.append(row[column_index - 1])
            
        return result

    def get_range(self, address: str) -> list[list]:
        """
        Возвращает значения из указанного диапазона ячеек.

        :param address: Диапазон, например, "A1:B2" или "A1".
        :return: Список списков значений (каждый вложенный список — строка).
        """
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId,
            range=address
        ).execute()

        values = result.get("values", [])
        return values

    def get_sheet(self, sheet_name: str) -> list[list]:
        """
        Возвращает все заполненные значения с указанного листа.

        :param sheet_name: Название листа (например, "Лист1").
        :return: Список списков, где каждый вложенный список — строка данных.
        """
        # Получаем данные с листа
        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheetId,
            range=sheet_name  # Указываем имя листа
        ).execute()

        # Возвращаем значения или пустой список, если данных нет
        return result.get("values", [])

    def set_cells(self, sheet_title: str, address: str, values: tuple[tuple[str]], majorDimension: Literal["ROWS", "COLUMNS"] = "ROWS"):
        """
        Устанавливает значение ячейкам в таблице
        
        :param sheet_title: Например, "Лист номер один".
        :param address: Например, "B2:D5".
        :param values: Список строк со значениями - [["B2", "C2", "D2"], ["B3", "C3", "D3"]].
        """
        return self.service.spreadsheets().values().batchUpdate(spreadsheetId = self.spreadsheetId, body = {
            # Данные воспринимаются, как вводимые пользователем (считается значение формул)
            "valueInputOption": "USER_ENTERED", 
            "data": [
                {
                    "range": f"{sheet_title}!{address}",
                    "majorDimension": majorDimension.upper(),  # Сначала заполнять строки, затем столбцы
                    "values": values
                }
            ]
        }).execute()
        
    def set_sheets(self, sheet_title: str, changes: set[int], sheet: list[list[str]]):
        """
        Устанавливает значения в указанных строках таблицы.

        :param sheet_title: Например, "Лист номер один".
        :param changes: Индексы строк (начиная с 0), которые нужно изменить.
        :param sheet: Двумерный список значений таблицы.
        """
        data = [
            {
                "range": f"{sheet_title}!A{index + 2}:Z{index + 2}",
                "majorDimension": "ROWS",
                "values": [sheet[index + 1]]
            }
            for index in sorted(changes)
        ]

        return self.service.spreadsheets().values().batchUpdate(
            spreadsheetId=self.spreadsheetId,
            body={
                "valueInputOption": "USER_ENTERED",
                "data": data
            }
        ).execute()
    
    def get_sheets(self) -> list[tuple[str, str]]:
        """
        Возвращает список листов в таблице.
        Возвращает список кортежей: (sheet_id, title)
        """
        spreadsheet = self.service.spreadsheets().get(spreadsheetId = self.spreadsheetId).execute()
        sheetList = spreadsheet.get("sheets")

        result = []
        for sheet in sheetList:
            result.append((sheet["properties"]["sheetId"], sheet["properties"]["title"]))
            
        # sheetId = sheetList[0]["properties"]["sheetId"]
        return result
    
    def add_sheet(self, title: str):
        """
        Добавление листа
        """
        return self.service.spreadsheets().batchUpdate(
            spreadsheetId = self.spreadsheetId,
            body = 
        {
        "requests": [
            {
            "addSheet": {
                "properties": {
                    "title": title,
                    "gridProperties": {
                        "rowCount": ROW_COUNT,
                        "columnCount": COLUMN_COUNT
                    }
                }
            }
            }
        ]
        }).execute()
    
    def get_comments(self, sheet_title: str, address: str) -> list[str]:
        """
        Возвращает комментарии для указанных ячеек в таблице.

        :param sheet_title: Название листа, например, "Лист номер один".
        :param address: Адрес ячейки или диапазона, например, "A1" или "B2:D5".

        :rtype: list
        :return: Список комментариев для указанных ячеек.
        """
        # Запрос к API для получения данных о ячейках
        range_ = f"{sheet_title}!{address}"
        result = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId, ranges=range_, includeGridData=True).execute()

        comments = []

        # Извлекаем комментарии из данных
        for sheet in result["sheets"]:
            for row in sheet["data"][0]["rowData"]:
                for cell in row.get("values", []):
                    if "note" in cell:
                        comments.append(cell["note"])
                    else:
                        comments.append("")

        return comments
    
    def get_all_comments(self, sheet_title: str) -> list[list[str]]:
        """
        Возвращает комментарии для всех ячеек на указанном листе.

        :param sheet_title: Название листа, например, "Лист номер один".

        :rtype: list[list[str]]
        :return: Двумерный список комментариев для всех ячеек листа.
        """
        # Запрос к API для получения данных о листе
        result = self.service.spreadsheets().get(
            spreadsheetId=self.spreadsheetId,
            ranges=sheet_title,
            includeGridData=True
        ).execute()

        comments = []

        # Извлекаем комментарии из данных
        for sheet in result.get("sheets", []):
            for row in sheet.get("data", [])[0].get("rowData", []):
                row_comments = []
                for cell in row.get("values", []):
                    row_comments.append(cell.get("note") or "")
                comments.append(row_comments)

        return comments

    def _column_letter_to_index(self, column: str) -> int:
        """
        Преобразует буквенное обозначение столбца (например, "A", "B", "AA") в индекс (0-based).

        :param column: Буква столбца, например, "A", "B", "AA".
        :return: Индекс столбца (0-based).
        """
        index = 0
        for char in column:
            index = index * 26 + (ord(char) - ord("A") + 1)

        return index - 1  # Делаем 0-based
    
    def _column_index_to_letter(self, index: int) -> str:
        """
        Преобразует числовой индекс столбца в буквенное обозначение (0 -> "A", 1 -> "B", ...).

        :param index: Индекс столбца (0-based).
        :return: Буквенное обозначение столбца, например, "A", "B", "C", ..., "AA", "AB".
        """
        result = ""
        while index >= 0:
            result = chr(index % 26 + 65) + result
            index = index // 26 - 1
        return result

    def _parse_cell_address(self, address: str) -> tuple[int, int]:
        """
        Парсит адрес ячейки (например, "B2") и возвращает индексы строки и столбца.

        :param address: Адрес ячейки, например, "B2".
        :return: (row_index, col_index), оба значения 0-based.
        """
        match = re.match(r"([A-Z]+)(\d+)", address)
        if not match:
            raise ValueError(f"Некорректный адрес ячейки: {address}")
        col, row = match.groups()
        return int(row) - 1, self._column_letter_to_index(col)

    def set_comment_cell(self, sheet_id: int, address: str, comment: str) -> None:
        """
        Устанавливает комментарий для указанной ячейки.

        :param sheet_id: Идентификатор листа (не порядковый!).
        :param address: Адрес ячейки, например, "A1".
        :param comment: Текст комментария, который нужно установить.
        """
        row_index, col_index = self._parse_cell_address(address)

        requests = [
            {
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_index,
                        "startColumnIndex": col_index,
                        "endRowIndex": row_index + 1,
                        "endColumnIndex": col_index + 1
                    },
                    "rows": [{"values": [{"note": comment}]}],
                    "fields": "note"
                }
            }
        ]

        return self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetId, body={"requests": requests}
        ).execute()

    def set_comments_row(self, sheet_id: int, cell_range: str, comments: list[str]) -> None:
        """
        Устанавливает комментарии для заданного диапазона ячеек в одной строке.
        Комментарии устанавливаются только в те ячейки, где соответствующий элемент списка comments не пустой.

        :param sheet_id: Идентификатор листа (не порядковый!).
        :param cell_range: Диапазон ячеек в строке, например "B2:D2".
        :param comments: Список комментариев, соответствующих каждой ячейке в диапазоне.
        """
        start_address, end_address = cell_range.split(":")
        start_row, start_col = self._parse_cell_address(start_address)
        end_row, end_col = self._parse_cell_address(end_address)

        if start_row != end_row:
            raise ValueError("Диапазон должен быть в одной строке.")

        if (end_col - start_col + 1) != len(comments):
            raise ValueError("Длина списка комментариев должна соответствовать количеству ячеек в диапазоне.")

        requests = []

        for i, comment in enumerate(comments):
            if comment:  # Только если комментарий не пустой
                col_index = start_col + i
                requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": start_row + 1,
                            "startColumnIndex": col_index,
                            "endColumnIndex": col_index + 1
                        },
                        "rows": [{"values": [{"note": comment}]}],
                        "fields": "note"
                    }
                })

        if requests:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId, body={"requests": requests}
            ).execute()
            
    def set_comments_sheet(self, sheet_id: int, comments: dict[str, str]) -> None:
        """
        Устанавливает комментарии для указанных диапазонов ячеек.

        :param sheet_id: Идентификатор листа (не порядковый!).
        :param comments: Словарь, где ключи - диапазоны (например, "A1" или "B2:C3"), 
                            а значения - комментарии, которые нужно установить.
        """
        requests = []

        for cell_range, comment in comments.items():
            start_address, *end_address = cell_range.split(":")
            start_row, start_col = self._parse_cell_address(start_address)
            end_row, end_col = self._parse_cell_address(end_address[0]) if end_address else (start_row, start_col)

            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row + 1,
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col + 1
                    },
                    "rows": [
                        {"values": [{"note": comment} for _ in range(start_col, end_col + 1)]}
                        for _ in range(start_row, end_row + 1)
                    ],
                    "fields": "note"
                }
            })

        if requests:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId, body={"requests": requests}
            ).execute()

    def set_background_color(self, sheet_title: str, color_map: dict[str, tuple]) -> None:
        """
        Устанавливает цвет фона для указанных ячеек.

        :param sheet_title: Название листа, например, "Лист1".
        :param color_map: Словарь, где ключ - адрес ячейки (например, "B1"), а значение - кортеж RGB (0-1), например, (0.8, 0.8, 0.8).
        """
        # Получаем sheetId по названию листа
        sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheet_id = None
        for sheet in sheet_metadata["sheets"]:
            if sheet["properties"]["title"] == sheet_title:
                sheet_id = sheet["properties"]["sheetId"]
                break

        if sheet_id is None:
            raise ValueError(f"Лист '{sheet_title}' не найден.")

        requests = []
        for address, color in color_map.items():
            match = re.match(r"([A-Z]+)(\d+)", address)
            if not match:
                raise ValueError(f"Некорректный адрес ячейки: {address}")

            col, row = match.groups()
            col_idx = self._column_letter_to_index(col)
            row_idx = int(row) - 1
            red, green, blue = color

            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": red,
                                "green": green,
                                "blue": blue
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })
        
        if requests:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId, body={"requests": requests}
            ).execute()
            
    def set_comments_and_colors(self, sheet_title: str, sheet_id: int, comments: dict[str, str], color_map: dict[str, tuple]) -> None:
        """
        Устанавливает комментарии и цвет фона для указанных ячеек на листе.

        :param sheet_title: Название листа, например, "Лист1".
        :param comments: Словарь, где ключ - адрес или диапазон (например, "A1" или "B2:C3"),
                            а значение - комментарий, который нужно установить.
        :param color_map: Словарь, где ключ - адрес ячейки (например, "B1"),
                            а значение - кортеж RGB (0-1), например, (0.8, 0.8, 0.8).
        """
        requests = []

        # Обработка комментариев
        for cell_range, comment in comments.items():
            start_address, *end_address = cell_range.split(":")
            start_row, start_col = self._parse_cell_address(start_address)
            end_row, end_col = self._parse_cell_address(end_address[0]) if end_address else (start_row, start_col)

            requests.append({
                "updateCells": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": start_row,
                        "endRowIndex": end_row + 1,
                        "startColumnIndex": start_col,
                        "endColumnIndex": end_col + 1
                    },
                    "rows": [
                        {"values": [{"note": comment} for _ in range(start_col, end_col + 1)]}
                        for _ in range(start_row, end_row + 1)
                    ],
                    "fields": "note"
                }
            })

        # Обработка цветов фона
        for address, color in color_map.items():
            match = re.match(r"([A-Z]+)(\d+)", address)
            if not match:
                raise ValueError(f"Некорректный адрес ячейки: {address}")

            col, row = match.groups()
            col_idx = self._column_letter_to_index(col)
            row_idx = int(row) - 1
            red, green, blue = color

            requests.append({
                "repeatCell": {
                    "range": {
                        "sheetId": sheet_id,
                        "startRowIndex": row_idx,
                        "endRowIndex": row_idx + 1,
                        "startColumnIndex": col_idx,
                        "endColumnIndex": col_idx + 1
                    },
                    "cell": {
                        "userEnteredFormat": {
                            "backgroundColor": {
                                "red": red,
                                "green": green,
                                "blue": blue
                            }
                        }
                    },
                    "fields": "userEnteredFormat.backgroundColor"
                }
            })

        if requests:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId, body={"requests": requests}
            ).execute()
            
    def update_sheet(self, sheet_id: int, values_map: dict[str, list[list[str]]] = None,
        comments: dict[str, str] = None, color_map: dict[str, tuple] = None,
        majorDimension: Literal["ROWS", "COLUMNS"] = "ROWS") -> list:
        """
        Обновляет значения, комментарии и цвета фона в таблице.

        :param sheet_title: Название листа, например, "Лист1".
        :param sheet_id: Идентификатор листа (не порядковый!).
        :param values_map: Словарь, где ключ - диапазон (например, "A1:B2"), 
                        а значение - список строк со значениями [["A1", "B1"], ["A2", "B2"]].
        :param comments: Словарь, где ключ - диапазон (например, "A1" или "B2:C3"), 
                        а значение - комментарий, который нужно установить.
        :param color_map: Словарь, где ключ - адрес ячейки (например, "B1"),
                        а значение - кортеж RGB (0-1), например, (0.8, 0.8, 0.8).
        :param majorDimension: Способ заполнения значений, "ROWS" или "COLUMNS" (по умолчанию "ROWS").
        """
        requests = []

        # Обновление значений ячеек
        if values_map:
            requests.extend([
                {
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": self._parse_cell_address(range_.split(":")[0])[0],
                            "endRowIndex": self._parse_cell_address(range_.split(":")[-1])[0] + 1,
                            "startColumnIndex": self._parse_cell_address(range_.split(":")[0])[1],
                            "endColumnIndex": self._parse_cell_address(range_.split(":")[-1])[1] + 1
                        },
                        "rows": [{"values": [{"userEnteredValue": {"stringValue": val}} for val in row]} for row in values],
                        "fields": "userEnteredValue"
                    }
                }
                for range_, values in values_map.items()
            ])

        # Обновление комментариев
        if comments:
            for cell_range, comment in comments.items():
                start_address, *end_address = cell_range.split(":")
                start_row, start_col = self._parse_cell_address(start_address)
                end_row, end_col = self._parse_cell_address(end_address[0]) if end_address else (start_row, start_col)

                requests.append({
                    "updateCells": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": start_row,
                            "endRowIndex": end_row + 1,
                            "startColumnIndex": start_col,
                            "endColumnIndex": end_col + 1
                        },
                        "rows": [
                            {"values": [{"note": comment} for _ in range(start_col, end_col + 1)]}
                            for _ in range(start_row, end_row + 1)
                        ],
                        "fields": "note"
                    }
                })

        # Обновление цвета фона
        if color_map:
            for address, color in color_map.items():
                match = re.match(r"([A-Z]+)(\d+)", address)
                if not match:
                    raise ValueError(f"Некорректный адрес ячейки: {address}")

                col, row = match.groups()
                col_idx = self._column_letter_to_index(col)
                row_idx = int(row) - 1
                red, green, blue = color

                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sheet_id,
                            "startRowIndex": row_idx,
                            "endRowIndex": row_idx + 1,
                            "startColumnIndex": col_idx,
                            "endColumnIndex": col_idx + 1
                        },
                        "cell": {
                            "userEnteredFormat": {
                                "backgroundColor": {
                                    "red": red,
                                    "green": green,
                                    "blue": blue
                                }
                            }
                        },
                        "fields": "userEnteredFormat.backgroundColor"
                    }
                })

        print(sheet_id, len(values_map), len(comments), len(color_map),"=>",len(requests))
        return requests
    
    def update_sheets(self, sheets: list):
        """
        :param sheets: Список листов `Sheet`, содержащих поля `id`, `cells_changes`, `comments_changes` и `colors_changes`.
        """
        
        requests = []
        for sheet in sheets:
            requests += self.update_sheet(sheet.id, sheet.cells_changes, sheet.comments_changes, sheet.colors_changes)
            
        print("total requests",len(requests))
        if requests:
            self.service.spreadsheets().batchUpdate(
                spreadsheetId=self.spreadsheetId, body={"requests": requests}
            ).execute()
    
    def create_sheet(self, title: str) -> int:
        """
        Добавляет новый лист в таблицу и возвращает его идентификатор.

        :param title: Название листа.
        :return: Идентификатор созданного листа.
        """
        response = self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetId,
            body={
                "requests": [
                    {
                        "addSheet": {
                            "properties": {
                                "title": title,
                                "gridProperties": {
                                    "rowCount": ROW_COUNT,
                                    "columnCount": COLUMN_COUNT
                                }
                            }
                        }
                    }
                ]
            }
        ).execute()

        # Возвращаем идентификатор созданного листа
        return response["replies"][0]["addSheet"]["properties"]["sheetId"]

    def delete_sheet(self, sheet_id: int):
        """
        Удаляет лист по его названию и возвращает его идентификатор.

        :param sheet_title: Название листа.
        :return: Идентификатор удалённого листа.
        """
        # Удаляем лист
        self.service.spreadsheets().batchUpdate(
            spreadsheetId=self.spreadsheetId,
            body={
                "requests": [
                    {
                        "deleteSheet": {
                            "sheetId": sheet_id
                        }
                    }
                ]
            }
        ).execute()

    def get_sheet_id(self, sheet_title: str) -> int | None:
        # Получаем идентификатор листа по его названию
        sheet_metadata = self.service.spreadsheets().get(spreadsheetId=self.spreadsheetId).execute()
        sheet_id = None
        for sheet in sheet_metadata["sheets"]:
            if sheet["properties"]["title"] == sheet_title:
                return sheet["properties"]["sheetId"]
                break

        return None