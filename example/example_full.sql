-- Create function that returns file
create or replace function excel_full() returns bytea language plpgsql as $$
declare 
  xls pgxls.xls; 
begin
  -- Create excel, first sheet has 1 column 30 wide and named "Styles" 
  xls := pgxls.create(array[30], sheet_name=>'Styles');
  -- Set value and style of cell through universal procedure
  call pgxls.set_cell_value(xls, 'Example full'::text, font_name=>'Times New Roman', font_size=>24, font_color=>'FF0000');
  -- Set value and style of cell through base procedures
  call pgxls.add_row(xls, 2);
  call pgxls.set_cell_timestamp(xls, now()::timestamp);
  call pgxls.set_cell_border(xls, 'thin');
  --
  call pgxls.add_row(xls, 2);
  call pgxls.set_cell_integer(xls, 123);  
  call pgxls.set_cell_alignment(xls, horizontal=>'left');
  --
  call pgxls.add_row(xls, 2);
  call pgxls.set_cell_numeric(xls, 1234567.89);  
  call pgxls.set_cell_format(xls, pgxls.get_format_code_numeric(3, true));
  call pgxls.set_cell_font(xls, name=>pgxls.font_name$monospace(), bold=>true);
  -- Add sheet named "Columns"
  call pgxls.add_sheet(xls, array[10,10,50], array['x','x/3','md5(x)'], 'Columns');
  -- Set format of column for numeric type 
  call pgxls.set_column_format_numeric(xls, 2, '0.000');
  -- Set alignment of column for all types
  call pgxls.set_column_alignment(xls, 3, horizontal=>'center');
  -- Set cell values, style defined by column and data type
  for x in 1..10 loop
    call pgxls.add_row(xls);
    call pgxls.set_cell_value(xls, x);
    call pgxls.set_cell_value(xls, x/3.0);
    call pgxls.set_cell_value(xls, md5(x::text));
  end loop;
    -- Add sheet by query
  call pgxls.add_sheet_by_query(xls, 'select * from pg_class order by 1 limit 10', 'Query');
  -- Return file  
  return pgxls.get_file(xls);      
end
$$;

-- Get file
select excel_full();
