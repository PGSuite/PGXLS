rem Examples of saving files in Windows command line

rem To convert hex to binary, use certutil utility and temporary file

rem When running psql on non-server, install postgres client and specify connection parameters via URI:
rem psql postgres://[USERNAME]:[PASSWORD]@[SERVER]/[DATABASE] -Aqt -c "...

rem 1. Create file by SQL query
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_class')" -o hex.tmp 
certutil -decodehex -f hex.tmp pg_class.xlsx

rem 2. Save Excel file on server by SQL query
psql -c "call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10')" 

rem 3. Save file from SQL function
psql -Aqt -c "select excel_top_relations_by_size()" -o hex.tmp 
certutil -decodehex -f hex.tmp top_relations_by_size.xlsx

