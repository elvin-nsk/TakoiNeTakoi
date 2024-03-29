# Takoi ili Ne Takoi

- По имени файла известен как `TakoiNeTakoi.gms`.
- Автор - **[Юрий Рождествин](https://vk.com/id172126342)**
- Проверенно работает в версиях **16, 24**.
- Язык: **русский**.
- Распространяется **бесплатно**, код **открытый**.
- **Поддерживается автором**.

## Установка

[Стандартная](https://github.com/elvin-nsk/cdr-vba/blob/master/articles/installation.md).

## Использование

Макрос ищет на странице, на слое или в выделенной области похожие или наоборот не похожие объекты.

В качестве критериев поиска выступают четыре параметра:
- тип фигуры: линия, текст, прямоугольник, эллипс, полигон
- цвет заполнения или отсутствие такового
- толщина линии обводки
- цвет линии обводки

Вы можете задать один или несколько критериев поиска включив соответствующий чекбокс. Задать параметр `Равно` или `Не равно`. Для толщины линии обводки есть еще два параметра `Больше` и `Меньше`.

Вам надо только выбрать исходный объект для поиска. Если выбрано для поиска несколько объектов, то параметры поиска берутся для "первого" объекта в выбранной области.

Есть три кнопки управления поиском:
- `Fix` - зафиксировать параметры поиска
- `Find` - искать объекты, соответствующие параметрам
- `Fix And Find` - искать в один клик. Те одним кликом и фиксируются параметры и происходит поиск.

Для нескольких параметров поиска действует логическая операция `AND`. То есть искомые объекты должны одновременно удовлетворять нескольким параметрам. При востребованности макроса добавлю операцию `OR`.