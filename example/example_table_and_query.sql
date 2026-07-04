-- Create function that returns file
create or replace function example.excel_table_and_query() returns bytea language plpgsql as $$
declare 
  xls pgxls.xls;
  rec record; 
begin
  -- Create excel document
  xls := pgxls.create();
  -- Add sheet by query
  call pgxls.add_sheet_by_query(xls, 'select * from pg_class order by 1', 'pg_class table');
  -- Formatting rows and cells is not possible because the data has already been output
  -- It is only possible to customize the current page.
  call pgxls.set_page_paper(xls, format=>'A5', orientation=>'landscape');  
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet 
  call pgxls.add_sheet(xls, array[10,15,50,15], array['OID','Schema','Table name','Owner'], 'Columns');
  --  Set format for new rows, current row with header not changed
  call pgxls.set_column_default_format_numeric(xls, 1, format_code=>pgxls.get_format_code_numeric(decimal_places=>0, thousands_separated=>true));
  call pgxls.set_column_default_format_numeric(xls, 4, font_size=>10);
  for rec in
      select oid,relnamespace::regnamespace as schema,
             relname                        as table_name,
             relowner::regrole::name       as owner
        from pg_class
        where relkind = 'r'
        order by 2,3    
  loop
    -- Add row
    call pgxls.add_row(xls);
    -- Set data from record into cells
    call pgxls.put_cell(xls, rec.oid);      
    call pgxls.put_cell(xls, rec.schema);   
    call pgxls.put_cell(xls, rec.table_name);
    call pgxls.put_cell(xls, rec.owner);
    --
    if rec.owner=current_user then
      call pgxls.format_row(xls, font_size=>12, font_bold=>true);
    end if;
  end loop;  
  -- Returns file (bytea type)
  return pgxls.get_file(xls);      
end
$$;
  
-- Get file
select example.excel_table_and_query();


