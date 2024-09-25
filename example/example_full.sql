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
  --------------------------------------------------------------------------------------------------------------------------------
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
  --------------------------------------------------------------------------------------------------------------------------------
   -- Add sheet by query
  call pgxls.add_sheet_by_query(xls, 'select * from pg_class order by 1 limit 10', 'SQL query');
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet with merged cells
  call pgxls.add_sheet(xls, array_fill(10, array[10]), null, 'Merge cells');
  for row in 1..10 loop
    call pgxls.add_row(xls);
    for col in 1..10 loop      
      call pgxls.set_cell_value(xls, row||','||col);
      if row=2 and col=2 then call pgxls.merge_cells(xls, 5/*column_count*/); call pgxls.set_cell_fill(xls, pgxls.color$light_red());   end if;
      if row=4 and col=4 then call pgxls.merge_cells(xls, row_count=>5);      call pgxls.set_cell_fill(xls, pgxls.color$light_green()); end if;
    end loop;    
    if row=6 then call pgxls.merge_cells(xls, 4, 4, 6); call pgxls.set_cell_fill(xls, pgxls.color$light_blue(), 6); end if;
  end loop;
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet with wrap text
  call pgxls.add_sheet(xls, array[40,30], sheet_name=>'Wrap text');
  call pgxls.set_column_alignment(xls, 1, horizontal=>'justify', text_wrap=>true);
  call pgxls.set_column_alignment(xls, 2, text_wrap=>true); 
  -- Text without line feed
  call pgxls.set_cell_text(xls, 'A database management system used to maintain relational databases is a relational database management system (RDBMS)');
  call pgxls.set_cell_value(xls, 'Row height calculated with assumptions and may be not optimal, usually line count is slightly larger'::text, font_size=>8); 
  -- Text with line feed <LF>
  call pgxls.add_row(xls,2);  
  call pgxls.set_cell_text(xls, 
    'PostgreSQL is a powerful, open source object-relational database system with over 35 years of active<LF>'|| chr(10) ||
    'development  that has earned it a strong reputation for reliability, feature robustness, and performance' 
  );
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet with print setup
  call pgxls.add_sheet(xls, array[15,15,15], sheet_name=>'Print setup');
  call pgxls.set_page_paper(xls, format=>'A5', orientation=>'landscape');
  call pgxls.set_page_header(xls, 'Example full / sheet "Print setup" page &P of &N');
  call pgxls.set_page_rows_repeat(xls, 3); -- table header on each page  
  call pgxls.set_cell_value(xls, 'Use "Print preview"'::text, font_size=>20, alignment_horizontal=>'center');
  call pgxls.merge_cells(xls, 3);
  call pgxls.add_row(xls);
  call pgxls.add_row(xls);
  call pgxls.set_column_border(xls, 1, 'thin');
  call pgxls.set_column_border(xls, 2, 'thin');
  call pgxls.set_column_border(xls, 3, 'thin');
  call pgxls.set_cell_value(xls, 'x'::text,  font_bold=>true, alignment_horizontal=>'center');
  call pgxls.set_cell_value(xls, '√x'::text, font_bold=>true, alignment_horizontal=>'center');
  call pgxls.set_cell_value(xls, 'x²'::text, font_bold=>true, alignment_horizontal=>'center');  
  call pgxls.set_column_format_numeric(xls, 2, '0.0000');
  for x in 1..100 loop
    call pgxls.add_row(xls);
    call pgxls.set_cell_value(xls, x);
    call pgxls.set_cell_value(xls, sqrt(x)::numeric);
    call pgxls.set_cell_value(xls, x*x);
  end loop;
  --------------------------------------------------------------------------------------------------------------------------------
  -- Return file  
  return pgxls.get_file(xls);      
end
$$;

-- Get file
select excel_full();
