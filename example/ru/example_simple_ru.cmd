rem Примеры сохранения файлов в командной строке Windows

rem Для преобразования hex в binary используется утилита certutil и промежуточный временный файл

rem При запуске psql не на сервере необходимо установить postgres client и указать параметры подключения через URI:
rem psql postgres://[ПОЛЬЗОВАТЕЛЬ]:[ПАРОЛЬ]@[СЕРВЕР]/[БАЗА_ДАННЫХ] -Aqt -c "...

rem  1. Создание файла по SQL-запросу
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_class')" -o hex.tmp 
certutil -decodehex -f hex.tmp pg_class.xlsx

rem  2. Сохранение файла из SQL-функции  
psql -Aqt -c "select excel_top_relations_by_size()" -o hex.tmp 
certutil -decodehex -f hex.tmp top_relations_by_size.xlsx

