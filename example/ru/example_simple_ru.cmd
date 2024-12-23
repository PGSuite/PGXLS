rem Примеры сохранения файлов в командной строке Windows

rem Для преобразования hex в binary используется утилита certutil и промежуточный временный файл

rem При запуске psql не на сервере необходимо установить postgres client и указать параметры подключения через URI:
rem psql postgres://[ПОЛЬЗОВАТЕЛЬ]:[ПАРОЛЬ]@[СЕРВЕР]/[БАЗА_ДАННЫХ] -Aqt -c "...

rem  1. Создание файла по SQL-запросу
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_tables')" -o hex.tmp 
certutil -decodehex -f hex.tmp pg_tables.xlsx

rem  2. Сохраняем Excel-файл на сервере по SQL-запросу
psql -c "call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10')" 

rem  3. Сохранение файла из SQL-функции  
psql -Aqt -c "select excel_top_relations_by_size()" -o hex.tmp 
certutil -decodehex -f hex.tmp top_relations_by_size.xlsx
