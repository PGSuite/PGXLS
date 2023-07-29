#!/bin/bash

# 1. Создание файла по SQL-запросу
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" | xxd -r -ps > pg_class.xlsx

# 2. Сохранение файла из SQL-функции  
psql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" | xxd -r -ps > top_relations_by_size.xlsx

