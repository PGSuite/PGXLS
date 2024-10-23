-- 1. Create and get file (bytea) by SQL query
select pgxls.get_file_by_query('select * from pg_tables');

-- 2. Save Excel file on server by SQL query
call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10');

-- 3. Create function that returns file (bytea)
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

