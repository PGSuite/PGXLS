rem Examples of saving files in Windows command line

rem To convert hex to binary, use certutil utility and temporary file

rem When running psql on non-server, install postgres client and specify connection parameters via URI:
rem psql postgres://[USERNAME]:[PASSWORD]@[SERVER]/[DATABASE] -Aqt -c "...

rem 1. Create file by SQL query
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" -o hex.tmp 
certutil -decodehex -f hex.tmp pg_class.xlsx

rem 2. Save file from SQL function
psql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" -o hex.tmp 
certutil -decodehex -f hex.tmp top_relations_by_size.xlsx

