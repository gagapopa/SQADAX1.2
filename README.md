SQADAX1.2
=========

САПР приложение для работы с FirebirdSQL
Пример кода для работы с БД SCADA фирмы ТЕКОН. Приложение обращается к БД FirebirdSQL, и БД реализованной в Excel.
Описание классов:

Основная часть реализована в mainForm супермакорониной, куда подключаются остальные классы по необходимости.
CPUlist по сути класс-модель данных, которую мы заполняем.
BDexelquerry обрабатывает данные из Excel формирует свою минимодель данных для передачи.
ExcelTableConn создает документацию из модели CPUlist

Используются Telerik.WinControls, которые выкладывать не стоит.
