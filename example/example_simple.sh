#!/bin/bash
# Examples of saving files in Linux command line

# When running psql on non-server, install postgres client and specify connection parameters via URI:
# psql postgres://[USERNAME]:[PASSWORD]@[SERVER]/[DATABASE] -Aqt -c "...

# 1. Create file by SQL query
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" | xxd -r -ps > pg_class.xlsx

# 2. Save file from SQL function
psql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" | xxd -r -ps > top_relations_by_size.xlsx

