-- Create function that returns file
create or replace function excel_full() returns bytea language plpgsql as $$
declare 
  xls pgxls.xls; 
begin
  -- Create excel, first sheet has 1 column 30 wide and named "Styles" 
  xls := pgxls.create(array[30], sheet_name=>'Styles');
  -- Set value and style of cell through universal procedure
  call pgxls.put_cell(xls, 'Example full'::text, font_name=>'Times New Roman', font_size=>24, font_color=>'FF0000');
  -- Set value and style of cell through base procedures
  call pgxls.add_row(xls, 2);
  call pgxls.put_cell_timestamp(xls, now()::timestamp);
  call pgxls.format_cell(xls, border_around=>'thin');
  --
  call pgxls.add_row(xls, 2);
  call pgxls.put_cell_integer(xls, 123);  
  call pgxls.format_cell(xls, alignment_horizontal=>'left');
  --
  call pgxls.add_row(xls, 2);
  call pgxls.put_cell_numeric(xls, 1234567.89);  
  call pgxls.format_cell(xls, format_code=>pgxls.get_format_code_numeric(3, true), font_name=>pgxls.font_name$monospace(), font_bold=>true);  
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet named "Columns"
  call pgxls.add_sheet(xls, array[10,15,50], array['x','x/3','md5(x)'], 'Columns');
  -- Set format of column for numeric type 
  call pgxls.set_column_format_numeric(xls, 2, format_code=>'0.0000', font_name=>pgxls.font_name$sans_serif());
  -- Set alignment of column for all types
  call pgxls.set_column_format(xls, 3, alignment_horizontal=>'center');
  -- Set cell values, style defined by column and data type
  for x in 1..10 loop
    call pgxls.add_row(xls);
    call pgxls.put_cell(xls, x);
    call pgxls.put_cell(xls, x/3.0);
    call pgxls.put_cell(xls, md5(x::text));
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
      call pgxls.put_cell(xls, row||','||col);
      if row=2 and col=2 then call pgxls.merge_cells(xls, 5/*column_count*/); call pgxls.format_cell(xls, fill_foreground_color=>pgxls.color$light_red());   end if;
      if row=4 and col=4 then call pgxls.merge_cells(xls, row_count=>5);      call pgxls.format_cell(xls, fill_foreground_color=>pgxls.color$light_green()); end if;
    end loop;    
    if row=6 then call pgxls.merge_cells(xls, 4, 4, 6); call pgxls.format_cell(xls, fill_foreground_color=>pgxls.color$light_blue(), column_=>6); end if;
  end loop;
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet with wrap text
  call pgxls.add_sheet(xls, array[40,30], name=>'Wrap text');
  call pgxls.set_column_format(xls, 1, alignment_horizontal=>'justify', alignment_text_wrap=>true);
  call pgxls.set_column_format(xls, 2, alignment_text_wrap=>true); 
  -- Text without line feed
  call pgxls.put_cell_text(xls, 'A database management system used to maintain relational databases is a relational database management system (RDBMS)');
  call pgxls.put_cell(xls, 'Row height calculated with assumptions and may be not optimal, usually line count is slightly larger'::text, font_size=>8); 
  -- Text with line feed <LF>
  call pgxls.add_row(xls,2);  
  call pgxls.put_cell_text(xls, 
    'PostgreSQL is a powerful, open source object-relational database system with over 35 years of active<LF>'|| chr(10) ||
    'development  that has earned it a strong reputation for reliability, feature robustness, and performance' 
  );
  --------------------------------------------------------------------------------------------------------------------------------
  -- Add sheet with print setup
  call pgxls.add_sheet(xls, array[10,15,10,60], name=>'Print setup');
  call pgxls.set_page_paper(xls, format=>'A5', orientation=>'landscape');
  call pgxls.set_page_header(xls, 'Example full / sheet "Print setup" page &P of &N');
  call pgxls.set_page_rows_repeat(xls, 3); -- table header on each page  
  call pgxls.put_cell(xls, 'Use "Print preview"'::text, font_size=>20, alignment_horizontal=>'center');
  call pgxls.merge_cells(xls, 4);
  call pgxls.add_row(xls);
  call pgxls.format_all(xls, border=>'thin');
  call pgxls.add_row(xls);  
  call pgxls.format_all(xls, fill_foreground_color=>pgxls.color$light_gray());
  call pgxls.put_cell(xls, 'x'::text);
  call pgxls.put_cell(xls, '√x'::text);
  call pgxls.put_cell(xls, 'x²'::text);  
  call pgxls.put_cell(xls, 'md5(x)'::text); 
  call pgxls.format_row(xls, alignment_horizontal=>'center', fill_foreground_color=>pgxls.color$dark_gray());
  call pgxls.set_column_format_numeric(xls, 2, format_code=>'0.0000');
  call pgxls.set_column_format(xls, 4, alignment_horizontal=>'center');
  for x in 1..100 loop
    call pgxls.add_row(xls);
    call pgxls.put_cell(xls, x);
    call pgxls.put_cell(xls, sqrt(x)::numeric);
    call pgxls.put_cell(xls, x*x);
    call pgxls.put_cell(xls, md5(x::text));
  end loop;
  --------------------------------------------------------------------------------------------------------------------------------
  -- Return file  
  return pgxls.get_file(xls);      
end
$$;

-- Get file
-- select excel_full();
