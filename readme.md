## PGXLS | Export to Excel from PostgreSQL

Tool `PGXLS` is schema with stored procedures for creating files (bytea type) in Excel format (.xlsx).\
Implemented dependence format on data type, conversion SQL query into sheet with autoformat and more.

### Installation ###

The installer is [pgxls_init.sql](https://github.com/PGXLS/PGXLS/blob/main/pgxls_init.sql) file that creates `pgxls` schema with the necessary procedures.  
The installation consists in executing it in the psql terminal client or SQL manager, for example:  

```bash
wget -O - https://raw.githubusercontent.com/PGXLS/PGXLS/main/pgxls_init.sql | psql -d [database]
```

If the developers are not superusers, need to grant them privileges on the `pgxls` schema and its procedures.  
To do this, use the SQL script [pgxls_grants.sql](https://github.com/PGXLS/PGXLS/blob/main/pgxls_grants.sql) with the roles variable. For example:

```bash
wget -O - https://raw.githubusercontent.com/PGXLS/PGXLS/main/pgxls_grants.sql | psql -d [database] -v roles=[developers]
```

[More info on site](https://pgxls.org/en/download/){:target="_blank"}
&nbsp;  
&nbsp;  

### Example #1 ###

Create and get file (bytea) by SQL query
```sql
select pgxls.get_file_by_query('select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10');
```

Save file on command line\
Linux
```bash
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" | xxd -r -ps > pg_class.xlsx
```

Windows
```cmd
psql -Aqt -c "select encode(pgxls.get_file_by_query('select * from pg_class'),'hex')" -o hex.tmp 
certutil -decodehex -f hex.tmp pg_class.xlsx
```

### Example #2 ###

```sql

-- Create function that returns file (bytea)
create or replace function excel_top_relations_by_size() returns bytea language plpgsql as $$
declare 
  rec record;
  xls pgxls.xls; 
begin
  -- Create document, specify widths and captions of columns in parameters
  xls := pgxls.create(array[10,80,15], array['oid','Name','Size, bytes']);
  -- Select data in loop
  for rec in
    select oid,relname,pg_relation_size(oid) size from pg_class order by 3 desc limit 10
  loop
    -- Add row
    call pgxls.add_row(xls);
    -- Set data from query into cells
    call pgxls.set_cell_value(xls, rec.oid); 
    call pgxls.set_cell_value(xls, rec.relname);
    call pgxls.set_cell_value(xls, rec.size);
  end loop;
  -- Returns file(bytea)
  return pgxls.get_file(xls); 
end
$$;
-- Get file
select excel_top_relations_by_size();

```

Save file on command line
```bash
sql -Aqt -c "select encode(excel_top_relations_by_size(),'hex')" | xxd -r -ps > top_relations_by_size.xlsx
```

[More examples on site](https://pgxls.org/en/#examples-full)
&nbsp;  
&nbsp;  

### Basic procedures ###
  
*   **pgxls.create** - create document
  
*   **pgxls.add_row** - add row
  
*   **pgxls.set_cell_value** - set cell value
  
*   **pgxls.get_file** - build and get file

[Documentation on site](https://pgxls.org/en/documentation/)
&nbsp;  
&nbsp;  


### Important qualities ### 

*   **Large files** - data row by row inserted into temporary table, which not requires memory. Separate function is implemented to get large file
*   **Auto-format** - for each column, format is configured depending on data type
*   **SQL queries** - it is possible to add sheet with the results of SQL query
*   **Styles** - for columns and cells, support format, font, border, fill and alignment
*   **Parallelism** - possible to create several files in parallel in one session
