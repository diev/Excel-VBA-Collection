# [Коллекция макросов Excel VBA](http://diev.github.io/Excel-VBA-Collection/)

Формат DBF, конечно, давно уже умер, но Федеральные службы РФ все еще охотно 
его используют. Например, ФСФМ (ФинМониторинг) требует отправлять ему файлы 
"в формате DBF", где 244 поля, и некоторые из них - типа C 254. Большинство 
древних утилиток для обработки DBF сходят с ума от таких объемов, а Excel 
никогда не сохранял требуемую структуру, более того - еще и даты часто 
коверкает. Так что даже открыть и посмотреть такой файл - целая проблема, 
а уж подкорректировать и сохранить как предписано - еще сложнее.

В итоге была написана эта программа для ручной обработки DBF, ни на кого не 
надеясь. Вернее, это была написана когда-то целая система Банк-Клиент на VBA 
Excel, и коллекция разных макросов из нее все еще на что-то годится, гибко 
обрабатывая типично русские превратности (типа суммы с разделителями любого 
вида, а не только того, что жестко задан в системе, да еще зависит от текущей 
языковой раскладки, как это бывает сделано у многих программеров, далеких от 
реальных работников).

![dbf161p2015.png](docs/assets/images/dbf161p2015.png)

На картинке выше обрабатывается файл в формате Приложения 4 "Структура файла 
передачи ОЭС" к Положению Банка России от 29 августа 2008 г. N 321-П 
"О порядке представления кредитными организациями в уполномоченный орган 
сведений, предусмотренных Федеральным законом "О противодействии легализации 
(отмыванию) доходов, полученных преступным путем, и финансированию 
терроризма". Помимо работы с таким форматом, программа также осуществляет 
контроль правильности заполнения полей по требованиям этого Положения 
несколькими способами.

Вы можете взять готовый бинарный файл XLSM с этой программой из Downloads
и при запуске обязательно разрешить макросы - только тогда появится меню
"Надстройки". Если боитесь запускать чужие бинарные файлы и макросы (и это 
правильно!) - открывайте редактор VBA в своем Excel (может понадобится в 
Настройках включить меню "Разработчик") и импортируйте туда прилагаемые 
исходные тексты (здесь они все в кодировке UTF-8).

## Как использовать

Раньше эта программа добавляла свою полосочку с кнопками меню, и все было 
замечательно. Затем Microsoft изобрела новый Ribbon, и пользовательское меню 
этой программы оказалась задвинута куда подальше - ищите в меню "Надстройки" - 
как это показано на скриншоте выше. И не забудьте разрешить макросы - иначе 
ничего не появится!

### Меню "Надстройки"

1. Загрузить - *загрузить из файла DBF, указанного в ячейке A1, или запросить 
его имя, если там пусто (текущее содержимое будет очищено)*
2. Добавить - *добавить из файла DBF, указанного в ячейке A1, или запросить 
его имя, если там пусто (текущее содержимое сохранится и будет дополнено)*
3. Просмотр - *просмотр всех полей на одной форме с индикацией ошибок*
4. Печать - *преобразовать выделенную строку в таблицу на отдельном листе 
для печати (одна из самых насущных функций и нравится проверяющим из ЦБ)*
5. Проверить - *проверить ячейку за ячейкой по заранее составленному перечню 
правил с индикацией нарушения и возможностью исправить*
6. Сохранить - *сохранить в файл DBF, указанный в ячейке A1, или запросить 
его имя, если там пусто (файл будет сформирован с той структурой, которая в 
строке 3 - описание см. ниже)*
7. Передать в Комиту - *сохранить в файл DBF и передать в папку для импорта 
в Комиту (специализированный АРМ Финмониторинга)*
8. Отправить в ЦБ - *сохранить в файл DBF и передать в папку для отправки 
на подпись, шифрование и далее в ПТК ПСД для отправки в ЦБ*

### Загрузка и сохранение

Если в ячейке A1 есть имя файла, при нажатии кнопки "Загрузить" - будет 
загружен именно этот файл. Если ячейка пуста - будет диалог выбора файла. 
При сохранении - аналогично. Файлы подразумеваются структуры DBF, хотя 
расширение их иное!
И никогда не сохраняйте DBF через меню самого Excel - он запишет в своей 
собственной структуре, подогнанной под текущие данные, но не в той, которая 
регламентирована.

### Работа со структурой

Одна из строк (третья на скриншоте) - структура загруженного DBF-файла, 
состоящая из столбцов-полей следующего вида:

`<Название поля> <Тип поля><Размер поля>`

Название отделяется от типа (поддерживаются C, D, L, N) пробелом, размер 
(в байтах) слитно с типом (у N может быть дробная часть после точки). 
Тип и размер - или Вы знаете, о чем речь, или это есть в документации по 
заполнению отчетности. Таким образом, Вы можете прочитать любой DBF-файл 
(без МЕМО), создать новый или сохранить с новой структурой.

Данные Вы можете копировать и вставлять какие угодно. Сохранение будет 
происходить в соответствии с указанной выше строкой описания структуры.

### Полезные мелочи ###

Попутно эта программа (и это основная ее нынешняя функция) проверяет данные 
по некоему набору логических правил, сильно облегчая жизнь отделу 
финмониторинга - даже в условиях существования других покупных монстров 
типа упомянутой Комиты, которые именно эти-то правила и пропускают мимо.

Все прочие исходные файлы, уже никак не относящиеся к этой задаче, убраны в 
папку **BClient** - на случай, если понадобится еще что-то из наработанного 
ранее.

Также там есть папка **Turniket**, где находится модуль подчистки грязных 
входных данных из разных источников СКУД и готовится финальная отчетная 
таблица с учетом отработанного времени сотрудниками для отдела кадров.

## Исходные тексты модулей

### Microsoft Excel Objects

* ЭтаКнига.cls - *всего две функции: добавить меню при загрузке Workbook и 
убрать его по ее закрытию*

### Forms

* UserForm1.frm - *форма полноэкранного просмотра записей*

### Modules

* Base36.bas - *работа с 36-ричными числами, популярными в Банке России*
* Bytes.bas - *работа с байтами - для CWinDos*
* ChkData.bas - *правила логического контроля настраивать здесь*
* CWinDos.bas - *ручная перекодировка 1251-866 с псевдографикой и фишками ЦБ*
* DBF3x.bas - *ручная работа в файлами DBF версии 3, загрузка и раскраска их, 
сохранение с заданной структурой*
* Export.bas - *пути экспорта*
* KeyValue.bas - *вычисления ключа счета, ИНН*
* Main.bas - *начальные инициализации классов, действия по завершении*
* MenuBar.bas - *пункты меню*
* MiscFiles.bas - *есть ли файл на диске, выбор файла и т.п.*
* MsgBoxes.bas - *разные красивые диалоги*
* Printf.bas - *аналог функции из языка Си*
* RuSumStr.bas - *чтение суммы в любом формате, сумма прописью - для платежек*
* SheetUtils.bas - *набросок макроса для сокращенной печати*
* StrFiles.bas - *работа с именами файловой системы*
* StrUtils.bas - *работа со строками*
* TextFile.bas - *работа с текстовыми файлами*

### Class Modules

* CApp.cls - *класс приложения с константами и параметрами*

## License

Licensed under the [Apache License, Version 2.0](LICENSE).
