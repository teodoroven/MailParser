import unittest
from modules import *
from main import *
from os import getenv
from dotenv import load_dotenv


class TestApp(unittest.TestCase):

    def setUp(self):
        Global.PRINT_ENABLED = True
        setup_logging()
        self.app: App = App()
        self.app.start(eel = False)

    def tearDown(self):
        self.app.stop()
    
    def test_get_accurate(self):
        string: str = "ООО «Рога и копыта, ко»"
        self.assertEqual(get_accurate(string, True), string)
        self.assertEqual(get_accurate(string, False), "ооо рога и копыта ко")
    
    def test_get_columns(self):
        self.assertEqual(get_columns(2), ["A", "B"])
    
    def test_get_max_date(self):
        self.assertEqual(get_max_date(["15.03.2025", "10.02.2025", "20.01.2025"]).strftime(DATE_FORMAT_2), "15.03.2025")
    
    def test_count_month(self):
        self.assertEqual(count_month("01.11.2024"), -1)
        self.assertEqual(count_month("01.12.2024"), 0)
        self.assertEqual(count_month("01.01.2025"), 1)
        self.assertEqual(count_month("01.02.2025"), 2)
        self.assertEqual(count_month("01.03.2025"), 3)

    def test_parse_year_2026(self):
        self.assertEqual(parse_year("01.12.", "01.01.2025"), "01.12.2024")
        self.assertEqual(parse_year("01.12.", "30.06.2025"), "01.12.2024")

        self.assertEqual(parse_year("01.12.", "01.07.2025"), "01.12.2025")
        self.assertEqual(parse_year("01.12.", "31.12.2025"), "01.12.2025")

        self.assertEqual(parse_year("01.12.", "19.01.2026"), "01.12.2025")
        self.assertEqual(parse_year("01.12.", "17.12.2025"), "01.12.2025")
        
        # первая половина 2026 → берём 2025
        self.assertEqual(parse_year("01.12.", "01.01.2026"), "01.12.2025")
        self.assertEqual(parse_year("01.12.", "30.06.2026"), "01.12.2025")

        # вторая половина 2026 → берём 2026
        self.assertEqual(parse_year("01.12.", "01.07.2026"), "01.12.2026")
        self.assertEqual(parse_year("01.12.", "31.12.2026"), "01.12.2026")

        
    def test_get_row(self):
        rows: list[str] = ["", "ООО Рога и Копыта", "ООО «Свин и ко»"]
        self.assertEqual(get_row(rows, "ООО 'Рога и Копыта'", False), 1)
        self.assertEqual(get_row(rows, "ООО Рога и Копыта", True), 1)
        self.assertEqual(get_row(rows, "ООО «Свин и ко»", True), 2)
        self.assertEqual(get_row(rows, "ооо свин и ко", False), 2)
        self.assertEqual(get_row(rows, "ООО «Свин и ко»", False), 2)
        self.assertEqual(get_row(rows, "ООО «Ромашка»", False), -1)
    
    def test_get_date(self):
        self.assertEqual(get_date(-11), "31.01.2024")
        self.assertEqual(get_date(-10), "29.02.2024")
        self.assertEqual(get_date(-9), "31.03.2024")
        self.assertEqual(get_date(-8), "30.04.2024")
        self.assertEqual(get_date(-7), "31.05.2024")
        self.assertEqual(get_date(-6), "30.06.2024")
        self.assertEqual(get_date(-5), "31.07.2024")
        self.assertEqual(get_date(-4), "31.08.2024")
        self.assertEqual(get_date(-3), "30.09.2024")
        self.assertEqual(get_date(-2), "31.10.2024")
        self.assertEqual(get_date(-1), "30.11.2024")
        self.assertEqual(get_date(0), "31.12.2024")
        self.assertEqual(get_date(1), "31.01.2025")
        self.assertEqual(get_date(2), "28.02.2025")
        self.assertEqual(get_date(3), "31.03.2025")
        self.assertEqual(get_date(4), "30.04.2025")
        self.assertEqual(get_date(5), "31.05.2025")
        self.assertEqual(get_date(6), "30.06.2025")
        self.assertEqual(get_date(7), "31.07.2025")
        self.assertEqual(get_date(8), "31.08.2025")
        self.assertEqual(get_date(9), "30.09.2025")
        self.assertEqual(get_date(10), "31.10.2025")
        self.assertEqual(get_date(11), "30.11.2025")
        self.assertEqual(get_date(12), "31.12.2025")
        self.assertEqual(get_date(13), "31.01.2026")
        self.assertEqual(get_date(14), "28.02.2026")
        self.assertEqual(get_date(15), "31.03.2026")
        self.assertEqual(get_date(-5, from_year=2024, from_month=6), "31.12.2023")
        self.assertEqual(get_date(-4, from_year=2024, from_month=6), "31.01.2024")
        self.assertEqual(get_date(-3, from_year=2024, from_month=6), "29.02.2024")
        self.assertEqual(get_date(-2, from_year=2024, from_month=6), "31.03.2024")
        self.assertEqual(get_date(-1, from_year=2024, from_month=6), "30.04.2024")
        self.assertEqual(get_date(0, from_year=2024, from_month=6), "31.05.2024")
        self.assertEqual(get_date(1, from_year=2024, from_month=6), "30.06.2024")
        self.assertEqual(get_date(2, from_year=2024, from_month=6), "31.07.2024")
        self.assertEqual(get_date(3, from_year=2024, from_month=6), "31.08.2024")
        self.assertEqual(get_date(4, from_year=2024, from_month=6), "30.09.2024")
        self.assertEqual(get_date(5, from_year=2024, from_month=6), "31.10.2024")
        self.assertEqual(get_date(6, from_year=2024, from_month=6), "30.11.2024")
        self.assertEqual(get_date(7, from_year=2024, from_month=6), "31.12.2024")
        self.assertEqual(get_date(8, from_year=2024, from_month=6), "31.01.2025")
        self.assertEqual(get_date(9, from_year=2024, from_month=6), "28.02.2025")
        self.assertEqual(get_date(10, from_year=2024, from_month=6), "31.03.2025")
        self.assertEqual(get_date(11, from_year=2024, from_month=6), "30.04.2025")
        self.assertEqual(get_date(12, from_year=2024, from_month=6), "31.05.2025")
    
    def test_get_messages(self):
        messages: list[Letter] = App().get_messages(False)
        
        i = 0
        for message in messages:
            print(i, message.subject, message.sender, message.date)
            i += 1
    
    def test_find_column(self):
        self.assertEqual(self.app.find_column("Лист 1", "22.03.2025"), 4)
        self.assertEqual(self.app.find_column("Лист 1", "22.07.2025"), 5)
        self.assertEqual(self.app.find_column("Лист 1", "22.04.2025"), 6)
    
    def test_get_update(self):
        self.assertEqual(self.app.get_update(
                Letter(b"1", 'Экономика по коммунальным ресурсам от 12.03.2025 ООО "Ромашка"', "example@example.example", datetime.now(), "")
            ), ("12.03.2025", 'ООО "Ромашка"')
        )
    
    def test_get_updates(self):
        self.assertEqual(self.app.get_updates({
                Letter(b"1", 'Экономика по коммунальным ресурсам от 12.03.2025 ООО "Ромашка"', "example@example.example", datetime.now(), ""): Trigger("subject", "includes", "Тест", True)
            }, True), ({'ООО "Ромашка"': ["12.03.2025"]}, {'ООО "Ромашка"': 'ООО "Ромашка"'})
        )
        self.assertEqual(self.app.get_updates({
                Letter(b"1", 'Экономика по коммунальным ресурсам от 12.03.2025 ООО "Ромашка"', "example@example.example", datetime.now(), ""): Trigger("subject", "includes", "Тест", 
            False)}, False), ({'ооо ромашка': ["12.03.2025"]}, {'ооо ромашка': 'ООО "Ромашка"'})
        )
        self.assertEqual(self.app.get_updates({
                Letter(b"1", 'Экономика по коммунальным ресурсам от 12.03.2025 ООО "Ромашка"', "example@example.example", datetime.now(), ""): Trigger("subject", "includes", "Тест", False),
                Letter(b"1", 'Экономика по коммунальным ресурсам от 13.03.2025 ООО Ромашка', "example@example.example", datetime.now(), ""): Trigger("subject", "includes", "Тест", False)
            }, False), ({'ооо ромашка': ["12.03.2025", "13.03.2025"]}, {'ооо ромашка': 'ООО "Ромашка"'})
        )
        
    def test_sync(self):
        self.app.disable_sync()
        self.app.sync(70, datetime(2025, 4, 30))
        
    def test_sync_small(self):
        self.app.disable_sync()
        self.app.sync(20, datetime(2025, 1, 19))
        
    def test_sync_large(self):
        self.app.disable_sync()
        self.app.run_eel()
        
        letters = [
            ('ип нефтьпрофи и к°', '18.06.2025'),
            ('ооо маркетингинвест', '18.08.2025'),
            ('тд itмир', '07.09.2025'),
            ('ип медменеджмент', '23.10.2025'),
            ('ао девелопментинвест и ко', '22.10.2025'),
            ('телекомстандарт и ко', '30.01.2026'),
            ('пао девелопменткапитал', '20.03.2026'),
            ('ооо энергоконтакт', '31.03.2025'),
            ('тд финансинвест групп', '02.07.2025'),
            ('автостандарт и к°', '05.05.2025'),
            ('ип девелопментальянс групп', '23.03.2026'),
            ('ао промменеджмент и ко', '12.01.2026'),
            ('ип газхолдинг групп', '08.08.2025'),
            ('гк ритейлхолдинг', '28.10.2025'),
            ('агросервис', '26.07.2025'),
            ('ао агроменеджмент групп', '24.08.2025'),
            ('зао ритейлмир', '04.04.2025'),
            ('пао мединвест', '30.03.2025'),
            ('ао девелопментвектор', '30.01.2025'),
            ('гк техноконтакт и к°', '22.01.2025'),
            ('ип маркетингтраст и партнеры', '25.01.2025'),
            ('пао телекоммир', '11.10.2025'),
            ('нефтьтрейд', '03.05.2025'),
            ('ао металлинвест', '11.09.2025'),
            ('нпо itкапитал', '05.02.2026'),
            ('тд автоменеджмент и ко', '11.06.2025'),
            ('зао логистиксоюз', '07.06.2025'),
            ('ооо маркетингцентр групп', '27.01.2026'),
            ('промсервис', '06.03.2025'),
            ('ип нефтькапитал', '28.05.2025'),
            ('пао техносервис групп', '06.01.2025'),
            ('ао финансконтакт', '21.01.2026'),
            ('тд ритейлтрейд и партнеры', '01.03.2025'),
            ('ао агрокапитал', '24.03.2025'),
            ('нпо газкапитал', '22.10.2025'),
            ('медкапитал и ко', '07.01.2026'),
            ('ао энергофактор', '28.06.2025'),
            ('ип девелопментвектор и ко', '21.06.2025'),
            ('нпо автоменеджмент', '21.04.2025'),
            ('финанстраст и партнеры', '02.03.2026'),
            ('ао энергоресурс', '20.07.2025'),
            ('itгарант и к°', '04.04.2025'),
            ('ип фармдиалог', '28.11.2025'),
            ('ооо финанстрейд', '20.01.2026'),
            ('девелопментпрофи и ко', '30.01.2026'),
            ('ип маркетингальянс', '01.11.2025'),
            ('гк ритейлинвест и к°', '21.02.2026'),
            ('ип медлидер', '22.02.2026'),
            ('зао энерготрейд', '17.10.2025'),
            ('пао металлинвест', '27.09.2025'),
            ('пао логистикконтакт', '30.11.2025'),
            ('ооо трансменеджмент и ко', '08.11.2025'),
            ('нпо трансрога', '20.02.2025'),
            ('гк автофонд и партнеры', '04.10.2025'),
            ('пао девелопментсервис', '01.01.2026'),
            ('нпо медальянс и ко', '22.12.2025'),
            ('гк стройгрупп и партнеры', '12.01.2026'),
            ('зао автоцентр', '23.02.2025'),
            ('пао логистикменеджмент и к°', '08.11.2025'),
            ('тд автоинвест', '31.10.2025'),
            ('ао автокопыта и к°', '08.08.2025'),
            ('пао ритейлтраст и партнеры', '02.04.2025'),
            ('гк энергоинвест', '28.08.2025'),
            ('itинвест и партнеры', '04.12.2025'),
            ('зао финансгрупп', '04.05.2025'),
            ('зао телекомфонд', '08.01.2025'),
            ('автоимпульс', '11.02.2026'),
            ('зао нефтьсервис', '23.05.2025'),
            ('гк медмир групп', '05.08.2025'),
            ('гк нефтьрога', '12.04.2025'),
            ('ооо торгтрейд', '27.04.2025'),
            ('ао нефтьсервис групп', '21.04.2025'),
            ('ао энергохолдинг', '29.05.2025'),
            ('ип энергогарант', '08.05.2025'),
            ('ао нефтьфонд', '24.02.2025'),
            ('нпо телекомгрупп групп', '11.02.2025'),
            ('тд стройхолдинг', '19.09.2025'),
            ('торгинвест и к°', '19.11.2025'),
            ('гк нефтьгарант', '07.08.2025', '26.08.2025'),
            ('логистикимпульс', '06.02.2026'),
            ('ип ритейлинвест', '07.02.2025'),
            ('ооо газсоюз', '18.03.2026'),
            ('нпо автоинвест', '14.01.2025'),
            ('зао торгресурс и к°', '27.02.2026'),
            ('ооо трансхолдинг', '03.07.2025'),
            ('ооо трансцентр', '04.11.2025'),
            ('тд эколидер', '30.01.2026'),
            ('зао металлстандарт и ко', '30.12.2025'),
            ('ао фармсоюз', '30.01.2025'),
            ('ао itинвест и к°', '13.03.2026'),
            ('ип энергодиалог и партнеры', '07.12.2025'),
            ('нпо стройменеджмент', '05.11.2025'),
            ('ооо маркетингфактор', '26.06.2025'),
            ('агрогрупп', '30.05.2025'),
            ('маркетингальянс', '30.01.2025'),
            ('ооо девелопментсервис и к°', '03.05.2025'),
            ('трансимпульс', '08.02.2026'),
            ('ооо ритейлкопыта', '30.09.2025'),
            ('пао ритейлфонд', '12.09.2025')
        ]
        
        result = self.app.sync()
        print("\n=================================")
        print(type(result),len(result) if result else 0)
        letters.sort()
        result.sort()
        
        if result != letters:
            for i in range(max(len(result),len(letters))):
                print(result[i],"\t",letters[i])
        else:
            print("SUCCESS!!!!!!!!!!!!!!!!!!")
            
        self.assertEqual(result, letters)
        
    def test_popup(self):
        self.app.html_filename = "popup.html"
        from debug import cmdline_args
        self.app.cmdline_args = cmdline_args
        self.app.run_eel()
        
        while self.app.check_working():
            sleep(1)
            
    def test_build(self):
        self.app.html_filename = "index.html"
        from debug import cmdline_args
        self.app.cmdline_args = cmdline_args
        self.app.run_eel()
        
        while self.app.check_working():
            sleep(1)
            
class TestSpreadsheet(unittest.TestCase):

    def setUp(self):
        setup_logging()
        self.app: App = App()
        self.app.start(eel = False)

    def tearDown(self):
        self.app.stop()

    def test_update_sheet(self):
        self.app.spreadsheet.update_sheet(0, {
            "A1": [["A1"]],
            "B2:C3": [["B2:C3", "B2:C3"], ["B2:C3", "B2:C3"]]
        }, {
            "A1": "Комментарий к A1",
            "B2:C3": "Комментарий к B2:C3",
        }, {
            "B1": (1, 0, 0),
            "D1": (1, 0, 0),
        })
        
    def test_set_comments_and_colors(self):
        self.app.spreadsheet.set_comments_and_colors("Лист 1", 0, {
            "A1": "Комментарий к A1",
            "B2:C3": "Комментарий к B2:C3",
        }, {
            "B1": (0.8, 0.8, 0.8),
            "D1": (0.8, 0.8, 0.8),
        })
        
    def test_set_comments_sheet(self):
        self.app.spreadsheet.set_comments_sheet(0, {
            "A1": "Комментарий к A1",
            "B2:C3": "Комментарий к B2:C3",
        })
        
    def test_get_comments(self):
        print("\n--------------------------------\n",
            self.app.spreadsheet.get_comments("Лист 1", "A9:D9"),
            "\n--------------------------------\n")
        
    def test_spreadsheet(self):
        print(self.app.spreadsheet.get("B1"))
        print(self.app.spreadsheet.get_row(1))
        print(self.app.spreadsheet.get_column("A"))
        print(self.app.spreadsheet.get_range("B2:D5"))
        
    def test_comment(self):
        comment: str = "test comment"
        address: str = "B1"
        sheet_title: str = "Лист 1"
        sheet_id: int = 0
        self.app.spreadsheet.set_comment(sheet_id, address, comment)
        self.assertEqual(self.app.spreadsheet.get_comments(sheet_title, address), [comment])
        self.app.spreadsheet.set_comment(sheet_id, address, "")
        self.assertEqual(self.app.spreadsheet.get_comments(sheet_title, address), [])
        
    def test_color(self):
        self.app.spreadsheet.set_background_color("Лист 1", {
            "B1": (0.8, 0.8, 0.8),
            "D1": (0.8, 0.8, 0.8),
        },)
        
    def test_create_row(self):
        print(self.app.create_row("Лист 1", "ля"))
        
    def test_create_columns(self):
        self.app.create_columns("Лист 1", 7)
    
    def test_create_sheet(self):
        spreadsheet: Spreadsheet = self.app.spreadsheet
        print(spreadsheet.create_sheet("Новый лист"))
        
    def test_delete_sheet(self):
        spreadsheet: Spreadsheet = self.app.spreadsheet
        spreadsheet.delete_sheet(0)
        
    def test_set_comments_row(self):
        spreadsheet: Spreadsheet = self.app.spreadsheet
        spreadsheet.set_comments_row(0, "B2:D2", ["Предупреждение", "", "Даты"])
        
class TestTask(unittest.TestCase):

    def setUp(self):
        setup_logging()
        pass

    def tearDown(self):
        pass

    def test_get_date(self):
        # Ежедневно в 12:00
        task: Task = Task(1, "days", "12:00")
        
        # В этот же день в 12:00
        self.assertEqual(task.get_date(datetime(2025, 3, 16, 11, 37)), datetime(2025, 3, 16, 12, 0))
        
        # На следующий день в 12:00
        self.assertEqual(task.get_date(datetime(2025, 3, 16, 16, 37)), datetime(2025, 3, 17, 12, 0))
        
        # На следующий день в 12:00
        self.assertEqual(task.get_date(datetime(2025, 3, 17, 12, 0)), datetime(2025, 3, 18, 12, 0))
        
        # Раз в два дня в 13:00
        task = Task(2, "days", "13:00")
        
        # В этот же день в 13:00
        self.assertEqual(task.get_date(datetime(2025, 3, 16, 11, 37)), datetime(2025, 3, 16, 13, 0))
        
        # Через день в 13:00
        self.assertEqual(task.get_date(datetime(2025, 3, 16, 16, 37)), datetime(2025, 3, 18, 13, 0))
        
        # Через в 13:00
        self.assertEqual(task.get_date(datetime(2025, 3, 17, 13, 0)), datetime(2025, 3, 19, 13, 0))
        
        # Каждые 30 секунд
        task = Task(30, "seconds")
        self.assertEqual(task.get_date(datetime(2025, 3, 16, 11, 36, 40)), datetime(2025, 3, 16, 11, 37, 10))
    
    def test_get_next_date(self):
        # Дата последней синхронизации (в прошлом)
        last_sync: datetime = datetime(2025, 3, 15, 11, 37, 50)
        
        # Ежедневно в 12:00
        task: Task = Task(1, "days", "12:00")
        self.assertEqual(get_next_date(task, last_sync), datetime(2025, 3, 17, 12, 00))

        # Раз в два дня в 13:00
        task = Task(2, "days", "13:00")
        self.assertEqual(get_next_date(task, last_sync), datetime(2025, 3, 17, 13, 00))

        # Каждые 30 секунд
        task = Task(30, "seconds")
        print(get_next_date(task, last_sync))

    def test_start_waiting(self):
        tasks: dict[Task, datetime] = {
            # Задача: дата следующей синхронизации
            Task(1, "days", "17:34"): datetime(2025, 3, 16, 12, 0),
            Task(30, "seconds"): datetime(2025, 3, 16, 17, 19, 28, 378015)
        }

        def update_tasks():
            for task, last_date in tasks.items():
                if last_date is None:
                    last_date = last_sync

                if last_date < datetime.now():
                    next_date: datetime = get_next_date(task, last_date)
                    tasks[task] = next_date
                
        def get_dates():
            return sorted(list(tasks.values()))

        def is_now(date: datetime):
            d1: datetime = datetime.now()
            d1 = d1.replace(microsecond=0)
            d2 = date.replace(microsecond=0)
            return d1 == d2

        in_progress: list[Task] = []

        last_sync: list[datetime] = [get_dates()[-1]]
        update_tasks()

        while True:
            if in_progress:
                s = "Выполняется синхронизация..."
            else:
                s = f"Следующая синхронизация будет {get_dates()[0]}"
            print("Последняя синхронизация была", last_sync[0],s)

            for task, date in tasks.items():
                if is_now(date):
                    if in_progress:
                        print("Пропущена синхронизация",datetime.now(),"во время выполнения другой синхронизации")
                        tasks[task] = get_next_date(task, last_sync[0])
                    else:
                        in_progress.append(task)
                        last_sync[0] = datetime.now()
                        
                        def target():
                            sleep(10)
                            tasks[task] = get_next_date(task, last_sync[0])
                            in_progress.clear()
                        Thread(target=target).start()

            sleep(1)

class TestMail(unittest.TestCase):

    def setUp(self):
        setup_logging()
        pass

    def tearDown(self):
        pass
    
    def generate_names(self, quantity = 250):
        import random
        from random import choice, sample, randint

        # Базовые слова для генерации
        FIRST_WORDS = ["ООО", "ЗАО", "ИП", "АО", "ПАО", "НПО", "ТД", "ГК", ""]
        COMPANY_TYPES = ["", "Групп", "Холдинг", "Инвест", "Трейд", "Сервис", "Менеджмент"]
        INDUSTRIES = [
            "Строй", "Техно", "Эко", "Агро", "Металл", "Нефть", "Газ", "Транс", 
            "Авто", "IT", "Мед", "Фарм", "Пром", "Торг", "Финанс", "Логистик", 
            "Энерго", "Телеком", "Ритейл", "Девелопмент", "Консалтинг", "Маркетинг"
        ]
        NOUNS = [
            "Мир", "Ресурс", "Гарант", "Профи", "Стандарт", "Капитал", "Траст", 
            "Лидер", "Фонд", "Партнер", "Сервис", "Центр", "Союз", "Альянс", 
            "Копыта", "Рога", "Вектор", "Контакт", "Фактор", "Диалог", "Импульс"
        ]
        SYMBOLS = ['"', "'", "«", ""]  # Пустая строка = кавычек нет

        def generate_company_core_name():
            """Генерирует ядро названия (без кавычек)"""
            first_word = choice(FIRST_WORDS)
            industry = choice(INDUSTRIES)
            noun = choice(NOUNS)
            company_type = choice(COMPANY_TYPES)
            
            # Собираем основное название
            if random.random() < 0.3:
                name = f"{industry}{noun}"
            else:
                name = f"{industry}{company_type or noun}"
            
            # Собираем полное название
            if first_word:
                full_name = f"{first_word} {name}"
            else:
                full_name = name
            
            # Иногда добавляем "и К°" / "и Партнеры" / "и Ко"
            if random.random() < 0.2:
                full_name += f" {choice(['и К°', 'и Партнеры', 'и Ко', 'Групп'])}"
            
            return full_name

        def add_random_quotes(name):
            """Добавляет случайные кавычки к названию"""
            if not name or random.random() < 0.5:  # 50% - без кавычек
                return name
            
            symbol = choice(SYMBOLS)
            if not symbol:  # если пустая строка
                return name
            
            if symbol == "«":
                return f"{symbol}{name}»"
            else:
                return f"{symbol}{name}{symbol}"

        # Генерируем 5000 уникальных **ядер** названий (без кавычек)
        unique_core_names = set()
        while len(unique_core_names) < quantity:
            unique_core_names.add(generate_company_core_name())

        # Добавляем случайные кавычки (но не меняем уникальность)
        return [add_random_quotes(name) for name in unique_core_names]
    
    def test_send_gmail(self):
        mail: IMAP = IMAP("imap.mail.ru", "sedoytankist@mail.ru", "zkiKq2mFTYiVYjUqer5Y")
        mail.send("sedoytankist@mail.ru", "Экономика по коммунальным ресурсам от 24.03.2025 ООО «Ромашка»", "Текст содержит текст")
        # mail: IMAP = IMAP("imap.mail.ru", "cedoy_tahkict@mail.ru", "Ast6w7SJZHbpkH7mA0rL")
    
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        me = "fedorfilipev2003@gmail.com"
        my_password = "xqmv athd wnhh yoqm"
        you = "sedoytankist@mail.ru"

        # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
        s = smtplib.SMTP_SSL('smtp.gmail.com')
        # uncomment if interested in the actual smtp conversation
        # s.set_debuglevel(1)
        # do the smtp auth; sends ehlo if it hasn't been sent already
        s.login(me, my_password)

        def send_gmail(to, subject, body):
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = me
            msg['To'] = you

            html = f'<html><body><p>{body}</p></body></html>'
            part2 = MIMEText(html, 'html')
            msg.attach(part2)
            
            s.sendmail(me, you, msg.as_string())
            
        send_gmail(you, "Экономика по коммунальным ресурсам от 08.07.2025 ТД МаркетингКопыта", "")
        # send_gmail(you, "Экономика по коммунальным ресурсам от 07.08.2025 ТД 'МаркетингКопыта'", "")
        # send_gmail(you, "Экономика по коммунальным ресурсам от 07.03.2025 ТД «МаркетингКопыта»", "")
        # send_gmail(you, "Экономика по коммунальным ресурсам от 31.12.2024 ТД МаркетингКопыта", "")
        # send_gmail(you, "Экономика по коммунальным ресурсам от 07.01.2025 ТД МаркетингКопыта", "")
        # send_gmail(you, "Экономика по коммунальным ресурсам от 07.09.2025 ТД МаркетингКопыта", "")
            
    def test_send_large(self):
        import smtplib
        from email.mime.multipart import MIMEMultipart
        from email.mime.text import MIMEText

        me = "fedorfilipev2003@gmail.com"
        my_password = "xqmv athd wnhh yoqm"
        you = "sedoytankist@mail.ru"

        # Send the message via gmail's regular server, over SSL - passwords are being sent, afterall
        s = smtplib.SMTP_SSL('smtp.gmail.com')
        # uncomment if interested in the actual smtp conversation
        # s.set_debuglevel(1)
        # do the smtp auth; sends ehlo if it hasn't been sent already
        s.login(me, my_password)

        def send_gmail(to, subject, body):
            msg = MIMEMultipart('alternative')
            msg['Subject'] = subject
            msg['From'] = me
            msg['To'] = you

            html = f'<html><body><p>{body}</p></body></html>'
            part2 = MIMEText(html, 'html')
            msg.attach(part2)
            
            s.sendmail(me, you, msg.as_string())

        max_messages = 500
        max_names = 150
        
        names: list[str] = self.generate_names(max_names)
        used_names: list[str] = []
        
        from random import randint
        from random import choice
        
        start_date = datetime(2025, 1, 1)
        end_date = datetime(2026, 3, 24)
        delta = end_date - start_date
        
        errors = []
        processing = []
        done = []
        all: dict[str, list[datetime]] = {}
        
        while len(done) < max_messages and len(all) < max_names:
            # def target():
            processing.append(True)
            noname = randint(1, 10000) < 50
            year = randint(1, 10000) < 5
            
            random_days: int = randint(0, delta.days)
            random_date: datetime = start_date + timedelta(days=random_days)
            if year:
                random_date.replace(year=2024)
            date = datetime.strftime(random_date, DATE_FORMAT_3)
            
            if noname:
                name = ""
            elif randint(1, 10000) < 7500 and used_names:
                name = choice(used_names)
            else:
                name = choice(names)
                used_names.append(name)
            
            subject: str = f"Экономика по коммунальным ресурсам от {date} {name}"
            
            try:
                pass
                send_gmail("sedoytankist@mail.ru", subject, "Текст содержит текст")
            except Exception as err:
                errors.append(err)
            else:
                key = get_accurate(name, False)
                
                if key in all:
                    all[key].append(random_date)
                else:
                    all[key] = [random_date]
            
            done.append(True)
            processing.remove(True)
            
            percent_messages: float = len(all) / max_messages
            percent_names: float = len(done) / max_messages
            if percent_messages > percent_names:
                percent = int(percent_messages*10000)/100
            else:
                percent = int(percent_names*10000)/100
            
            print(f"done: {percent}%",end=" "*70+"\r")
            
            # while len(processing) > 20:
            #     sleep(0.01)
            
            # Thread(target=target).start()
            
        # def target():
        #     input(">>>")
        #     while len(done) < test:
        #         done.append(True)
                
        # Thread(target=target).start()
        
        # while len(done) < test:
        #     sleep(1)
        
        result = ""
        for key, dates in all.items():
            dates.sort()
            
            s = f"{key}:", ";".join([date.strftime(DATE_FORMAT_3) for date in dates])
            result += f"{s}\n"
        
        with open("filename.txt","w",encoding="utf-8") as file:
            file.write(result)
            
        if errors:
            print(errors[0])
            
        s.quit()


class Test(unittest.TestCase):

    def setUp(self):
        Global.PRINT_ENABLED = True
        setup_logging()

    def tearDown(self):
        pass
    
    def test_encrypt(self):
        load_dotenv()
        KEY = base64.b64decode(os.environ["KEY_B64"])
        IV  = base64.b64decode(os.environ["IV_B64"])
        key, iv = load_keys()

        for filename in Data(None).filenames:
            open_filename = join("DATA", filename)
            encrypted_filename = filename

            if not isfile(open_filename):
                raise FileNotFoundError(f"Файл не найден: {open_filename}")

            with open(open_filename, "rb") as f:
                encrypted = f.read()

            decrypted = decrypt(encrypted, KEY, IV)
            encrypted = encrypt(decrypted, key, iv)

            # записываем как DATA1 / DATA2
            with open(encrypted_filename, "wb") as f:
                f.write(encrypted)
    
    def test_parse_column(self):
        # Базовая дата письма для тестов
        letter_date = datetime(2025, 7, 1).strftime(DATE_FORMAT_3)

        # --- Месяцы ---
        self.assertEqual(parse_column("Декабрь 2024", letter_date), "01.12.2024")
        self.assertEqual(parse_column("Январь 2025", letter_date), "01.01.2025")
        self.assertEqual(parse_column("Февраль 2025", letter_date), "01.02.2025")
        self.assertEqual(parse_column("март 2025", letter_date), "01.03.2025")
        self.assertEqual(parse_column(" апрель 2025", letter_date), "01.04.2025")
        self.assertEqual(parse_column("МАЙ 2025 ", letter_date), "01.05.2025")
        self.assertEqual(parse_column("Июнь 2025", letter_date), "01.06.2025")
        self.assertEqual(parse_column("Июль 2025", letter_date), "01.07.2025")
        self.assertEqual(parse_column("Август 2025", letter_date), "01.08.2025")
        self.assertEqual(parse_column("Сентябрь 2025", letter_date), "01.09.2025")
        self.assertEqual(parse_column("Октябрь 2025", letter_date), "01.10.2025")
        self.assertEqual(parse_column("Ноябрь 2025", letter_date), "01.11.2025")
        self.assertEqual(parse_column("Декабрь 2025", letter_date), "01.12.2025")
        self.assertEqual(parse_column("Январь 2026", letter_date), "01.01.2026")

        # --- Кварталы ---
        # Для письма 01.07.2025
        self.assertEqual(parse_column("1 квартал", datetime(2025, 1, 15).strftime(DATE_FORMAT_3)), "01.03.2024")
        self.assertEqual(parse_column("1 квартал", datetime(2025, 4, 15).strftime(DATE_FORMAT_3)), "01.03.2025")
        self.assertEqual(parse_column("2 квартал", datetime(2025, 2, 1).strftime(DATE_FORMAT_3)), "01.06.2024")
        self.assertEqual(parse_column("2 квартал", datetime(2025, 7, 1).strftime(DATE_FORMAT_3)), "01.06.2025")
        self.assertEqual(parse_column("3 квартал", datetime(2025, 1, 10).strftime(DATE_FORMAT_3)), "01.09.2024")
        self.assertEqual(parse_column("3 квартал", datetime(2025, 10, 10).strftime(DATE_FORMAT_3)), "01.09.2025")
        self.assertEqual(parse_column("4 квартал", datetime(2025, 11, 22).strftime(DATE_FORMAT_3)), "01.12.2024")
        self.assertEqual(parse_column("4 квартал", datetime(2025, 12, 1).strftime(DATE_FORMAT_3)), "01.12.2025")
        self.assertEqual(parse_column("4 квартал", datetime(2026, 1, 19).strftime(DATE_FORMAT_3)), "01.12.2025")

        # --- Полугодие ---
        self.assertEqual(parse_column("полугодие", datetime(2025, 7, 5).strftime(DATE_FORMAT_3)), "01.06.2025")
        self.assertEqual(parse_column("полугодие", datetime(2025, 6, 30).strftime(DATE_FORMAT_3)), "01.06.2025")
        self.assertEqual(parse_column("полугодие", datetime(2025, 1, 10).strftime(DATE_FORMAT_3)), "01.06.2024")

        # --- Год ---
        self.assertEqual(parse_column("год", datetime(2025, 12, 20).strftime(DATE_FORMAT_3)), "01.12.2025")
        self.assertEqual(parse_column("год", datetime(2026, 1, 19).strftime(DATE_FORMAT_3)), "01.12.2025")
        self.assertEqual(parse_column("год", datetime(2025, 1, 5).strftime(DATE_FORMAT_3)), "01.12.2024")

        # --- Месяцы N месяцев ---
        self.assertEqual(parse_column("5 месяцев", datetime(2025, 7, 5).strftime(DATE_FORMAT_3)), "01.05.2025")
        self.assertEqual(parse_column("9 месяцев", datetime(2025, 10, 10).strftime(DATE_FORMAT_3)), "01.09.2025")


# Точка входа для тестирования
if __name__ == "__main__":
    unittest.main()
    # TestMail().test_send_large()