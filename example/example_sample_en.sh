#!/bin/bash

# 1. Create file by SQL query
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" | xxd -r -ps > pg_class.xlsx

# 2. Save file from SQL function
psql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" | xxd -r -ps > top_relations_by_size.xlsx

