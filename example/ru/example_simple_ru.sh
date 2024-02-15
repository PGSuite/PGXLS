#!/bin/bash
# Примеры сохранения файлов в командной строке Linux
#
# При запуске psql не на сервере необходимо установить postgres client и указать параметры подключения через URI:
# psql postgres://[ПОЛЬЗОВАТЕЛЬ]:[ПАРОЛЬ]@[СЕРВЕР]/[БАЗА_ДАННЫХ] -Aqt -c "...

# 1. Создание файла по SQL-запросу
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" | xxd -r -ps > pg_class.xlsx

# 2. Сохранение файла из SQL-функции  
psql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" | xxd -r -ps > top_relations_by_size.xlsx

