EasyExcel
====

Библиотека создана для упрощения работы с excel.

Зависимости:
----
* Excel 2019 (older versions are not tested)

Считывание данных
----

Q: Документ не редактируется после взаимодействия с данной библиотекой.
A: Значит в результате программы не был вызван метод ```elements.Stop();```, вследствии чего копия Excel всё ещё удерживает документ. Чтобы это исправить надо зайти в ```Диспетчер задач``` и остановить ```фоновые``` процессы связянные с Excel. Такие как ```Microsoft Excel```.

#### Чтение всей страницы

	using EasyExcel;

	Elements elements = new Elements();
	Reader reader = new Reader(ref elements);
	
	// Запускает Excel в скрытом режиме
	elements.Start();
	// Открываем документ (можно писать без расширения ("...\data"))
	elements.Open(@"C:\Users\User\Documents\data.xlsx");
	// Перемещаемся на первую страницу
        elements.setWorksheet(1);
	// Считываем данные
        object[,] result = reader.Read();
	// Останавливаем Excel
	elements.Stop();

#### Чтение части данных

	using EasyExcel;
	
	Elements elements = new Elements();
	Reader reader = new Reader(ref elements);
	
	// Запускает Excel в скрытом режиме
	elements.Start();
	// Открываем документ (можно писать без расширения ("...\data"))
	elements.Open(@"C:\Users\User\Documents\data.xlsx");
	// Перемещаемся на первую страницу
        elements.setWorksheet(1);
	// Создаем точку, с которой начнется считывание данных
	Point point = reader.createPoint();
	// Устанавливаем координаты x, y
        point.set(2,2);
	// Считываем данные
        object[,] result = reader.Read(point);
	// Останавливаем Excel
	elements.Stop();

#### Чтение определенного диапазона данных

	using EasyExcel;
	
	Elements elements = new Elements();
	Reader reader = new Reader(ref elements);
	
	// Запускает Excel в скрытом режиме
	elements.Start();
	// Открываем документ (можно писать без расширения ("...\data"))
	elements.Open(@"C:\Users\User\Documents\data.xlsx");
	// Перемещаемся на первую страницу
        elements.setWorksheet(1);
	// Создаем точки, в диапазоне которых начнется считывание данных
	Point point1 = reader.createPoint();
	Point point2 = reader.createPoint();
	// Устанавливаем координаты x, y
        point1.set(2,2);
        point2.set(2,2);

	// Считываем данные
        object[,] result = reader.Read(point1, point2);
	// Останавливаем Excel
	elements.Stop();

Запись данных
----

#### Создание документа

	using EasyExcel;

	Elements elements = new Elements();
        Writer writer= new Writer(ref elements);

	// Запускает Excel в скрытом режиме
	elements.Start();
	// Создаем рабочую книгу
	elements.createWorkbook();
	// Сохраняем документ (можно писать без расширения ("...\data"))
	elements.Save(@"C:\Users\User\Documents\data.xlsx");
	// Останавливаем Excel
	elements.Stop();

#### Записывание данных, начиная с первой клетки

	using EasyExcel;

	Elements elements = new Elements();
        Writer writer= new Writer(ref elements);

	// Данные для записи
	 object[,] data = new object[,] { { "Name", "Second Name", "Account id" },
                				{ "MX", "Name2", 999 }};

	// Запускает Excel в скрытом режиме
	elements.Start();
	// Устанавливаем рабочий лист
	elements.setWorksheet(1);
	// Записываем данные
	writer.Write(data);
	// Сохраняем документ (можно писать без расширения ("...\data"))
	elements.Save(@"C:\Users\User\Documents\data.xlsx");
	// Останавливаем Excel
	elements.Stop();




#### Записывание данных, начиная с указанной клетки

	using EasyExcel;

	Elements elements = new Elements();
        Writer writer= new Writer(ref elements);

	// Данные для записи
	 object[,] data = new object[,] { { "Name", "Second Name", "Account id" },
                				{ "MX", "Name2", 999 }};

	// Запускает Excel в скрытом режиме
	elements.Start();
	// Устанавливаем рабочий лист
	elements.setWorksheet(1);
	
	// Создаем точку, с которой начнется запись данных
	Point point = writer.createPoint();
	// Устанавливаем координаты x, y
        point1.set(10,15);

	// Записываем данные
	writer.Write(data, point1);
	// Сохраняем документ (можно писать без расширения ("...\data"))
	elements.Save(@"C:\Users\User\Documents\data.xlsx");
	// Останавливаем Excel
	elements.Stop();

Редактирование данных
----
	
	using EasyExcel;

	Elements elements = new Elements();
        Writer writer= new Writer(ref elements);

	// Данные для записи
	 object[,] data = new object[,] { { "Name", "Second Name", "Account id" },
                				{ "MX", "Name2", 999 }};

	// Запускает Excel в скрытом режиме
	elements.Start();
	// Открываем документ (можно писать без расширения ("...\data"))
	elements.Open(@"C:\Users\User\Documents\data.xlsx");
	// Устанавливаем рабочий лист
	elements.setWorksheet(1);
	
	// Создаем точку, с которой начнется запись данных
	Point point = writer.createPoint();
	// Устанавливаем координаты x, y
        point1.set(10,15);

	// Записываем данные
	writer.Write(data, point1);
	// Сохраняем документ (можно писать без расширения ("...\data"))
	elements.Save(@"C:\Users\User\Documents\data.xlsx");
	// Останавливаем Excel
	elements.Stop();

























PS: I hope you are not using excel instead of a database.