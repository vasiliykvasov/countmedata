# countmedata
VBA макрос для Excel, который считает для меня данные по сценарию в первой ячейке столбца

## Приступая к работе
Для запуска макроса нужно скачать его и импортировать файл макроса в редактор Visual Basic в приложении Microsoft Excel. Для работы с макросами должна быть активна вкладка Разработчик.

## Для чего этот макрос
Я использую этот макрос для автоматических вычислений значений из несгруппированных ячеек рядов, деления одних вычисленных значений на другие, а также очищения ячеек при создании отчетов в Microsoft Excel.

## Как работает макрос
Макрос работает со столбцами выделенного диапазона ячеек. Под каждым выделенным стролбцом он записывает результат в виде частного ячеек, которые расположены под выделенным диапазоном, формулы с несгруппированными ячейками текущего ряда или очищает данные под колонкой.

## Как запустить макрос
Для работы макроса нужно вписать в ячейки 1 строки, какие действия необходимо произвести с соответствующим столбцом. В данный момент макрос работает со следующими значениями в 1 ячейках:
+ *Z:A* - записывает в ячейке под последней строкой выделенного диапазона в соответствующем столбце формулу '=$Z$*0*/$A$*0*', где вместо нуля номер строки под выделенным диапазоном.
+ *ОЧИСТИТЬ* - очищает ячейку под последней строкой выделенного диапазона в соответствующем столбце.
+ СУММ - записывает в ячейке под последней строкой выделенного диапазона в соответствующем столбце формулу '=СУММ(*несгруппированные ячейки соответствующего ряда*)'.
+ СРЗНАЧ - ззаписывает в ячейке под последней строкой выделенного диапазона в соответствующем столбце формулу '=СРЗНАЧ(*несгруппированные ячейки соответствующего ряда*)'.

1. Заполните ряд 1 ячейки нужных стобцов для определения режима работы макроса.
2. Выделите данные, с которыми будет работать макрос.
3. Запустите макрос.
4. Под последней строкой выделенного диапазона в каждом ряду появится результат работы макроса.

## Управление версии
Для управления версиями я использую Git.

## Автор
+ Василий Квасов

## Лицензия
Этот проект лицензируется в соответствии с лицензией Beerware
