#!/bin/bash
# Примеры сохранения файлов в командной строке Linux
#
# Для преобразования hex в binary используется утилита xxd
#
# При запуске psql не на сервере необходимо установить postgres client и указать параметры подключения через URI:
# psql postgres://[ПОЛЬЗОВАТЕЛЬ]:[ПАРОЛЬ]@[СЕРВЕР]/[БАЗА_ДАННЫХ] -Aqt -c "...

# 1. Создание файла по SQL-запросу
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_tables')" | xxd -r -ps > pg_tables.xlsx

# 2. Сохраняем Excel-файл на сервере по SQL-запросу
psql -c "call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10')" 

# 3. Сохранение файла из SQL-функции  
psql -Aqt -c "select excel_top_relations_by_size()" | xxd -r -ps > top_relations_by_size.xlsx

