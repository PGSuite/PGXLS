## PGXLS | Export to Excel from PostgreSQL

Tool `PGXLS` is schema with stored procedures for creating files (bytea type) in Excel format (.xlsx).  
Implemented dependence format on data type, conversion SQL query into sheet with autoformat and more.

### Installation ###

The installer is [pgxls_init.sql](https://github.com/PGSuite/PGXLS/raw/main/pgxls_init.sql) file that creates `pgxls` schema with the necessary procedures.  
The installation consists in executing it in the psql terminal client or SQL manager, for example:  

```bash
wget -O - -q https://github.com/PGSuite/PGXLS/raw/main/pgxls_init.sql | psql -d [database]
```
### Extension ###

To install `PGXLS` as an extension, unpack the [pgxls_extension.tar](https://github.com/PGSuite/PGXLS/raw/main/extension/pgxls_extension.tar) archive into the [sharedir]/extension directory of the postgres installation, for example (run as root):

```bash
wget -O - -q https://github.com/PGSuite/PGXLS/raw/main/extension/pgxls_extension.tar | tar x -C `su - postgres -c "pg_config --sharedir"`/extension
```
More info on page [download](https://pgxls.org/en/download/)

### Example #1 ###

Create and get file (bytea) by SQL query
```sql
select pgxls.get_file_by_query('select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10');
```

Save file on command line  
```bash
psql -Aqt -c "select pgxls.get_file_by_query('select * from pg_class')" | xxd -r -ps > pg_class.xlsx
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
sql -Aqt -c "select excel_top_relations_by_size()" | xxd -r -ps > top_relations_by_size.xlsx
```

All examples in directory [example](https://github.com/PGSuite/PGXLS/tree/main/example)

### Basic procedures ###
  
*   **pgxls.create** - create document
  
*   **pgxls.add_row** - add row
  
*   **pgxls.set_cell_value** - set cell value
  
*   **pgxls.get_file** - build and get file

*   **pgxls.get_file_by_query**	- create file by SQL query (wrapper)


Documentation in file [documentation/documentation.html](https://htmlpreview.github.io/?https://github.com/PGSuite/PGXLS/blob/main/documentation/documentation.html)


### Important qualities ### 

*   **Large files** - data row by row inserted into temporary table, which not requires memory. Separate function is implemented to get large file
*   **Auto-format** - for each column, format is configured depending on data type
*   **SQL queries** - it is possible to add sheet with the results of SQL query
*   **Styles**      - for columns and cells, support format, font, border, fill and alignment
*   **Print setup** - for each sheet, the paper format and orientation, title, repeating rows (table header) and margins are specified
*   **Parallelism** - possible to create several files in parallel in one session

Overview on site [pgxls.org](https://pgxls.org/en/)

### Support ### 

Of course you can create an issue, I will answer all requests.  
Also I will help to install and use the tool.  
Welcome to discussion !  

WhatsApp: [PGSuite (+7-936-1397626)](https://wa.me/79361397626)  
email: [support\@pgsuite.org](mailto:support@pgsuite.org?subject=PGXLS)

