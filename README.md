


## История

<br>


<p style="color: gray;">26.07.2024 16:56:38</p>

<font color="purple">Обновление v2.1:</font>


- Оптимизация make_xls.py
  - Теперь если уже создан файл summary.xlsx, то он не будет пересоздаваться заново, будут просто перезаписаны значения 
  - Убрано выделение всех клеток со значением, что занимало много времени

- Оптимизация sql_requests.py
  - Если выбрана больше половины форм, то будет использоваться конструкция NOT IN в запросе

<br>
<br>


<p style="color: gray;">19.07.2024 17:35:17</p>

Обновление v2.0 
-

- Оптимизация sql_requests.py
  - Убрана функция создающая дапазоны, для использования конструкций OR и BETWEEN, вместо этого в запросе используется оператор IN
  - Изменен порядок условий
  - Убран лишний код

<br>
<br>


<p style="color: gray;">14.07.2024 16:40:47</p>

Релиз
-

- Релиз программы

