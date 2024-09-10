#!/bin/bash
# Примеры сохранения файлов в командной строке Linux
#
# Для преобразования hex в binary используется утилита xxd
#
# При запуске psql не на сервере необходимо установить postgres client и указать параметры подключения через URI:
# psql postgres://[ПОЛЬЗОВАТЕЛЬ]:[ПАРОЛЬ]@[СЕРВЕР]/[БАЗА_ДАННЫХ] -Aqt -c "...

# 1. Создание файла по SQL-запросу
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_class')" | xxd -r -ps > pg_class.xlsx

# 2. Сохранение файла из SQL-функции  
psql -Aqt -c "select excel_top_relations_by_size()" | xxd -r -ps > top_relations_by_size.xlsx

