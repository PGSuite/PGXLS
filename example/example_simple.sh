#!/bin/bash
# Examples of saving files in Linux command line

# To convert hex to binary, use xxd utility

# When running psql on non-server, install postgres client and specify connection parameters via URI:
# psql postgres://[USERNAME]:[PASSWORD]@[SERVER]/[DATABASE] -Aqt -c "...

# 1. Create file by SQL query
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_class')" | xxd -r -ps > pg_class.xlsx

# 2. Save Excel file on server by SQL query
psql -c "call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10')" 

# 3. Save file from SQL function
psql -Aqt -c "select excel_top_relations_by_size()" | xxd -r -ps > top_relations_by_size.xlsx

