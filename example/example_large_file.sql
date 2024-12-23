-- Create function that returns large file by parts
create or replace function excel_large_file() returns setof bytea language plpgsql as $$
declare 
  xls pgxls.xls; 
  v_value bigint;
begin
  xls := pgxls.create(array[80]);
  call pgxls.put_cell(xls, 'Large file'::text, font_bold=>true, font_size=>32, alignment_horizontal=>'center');
  call pgxls.add_row(xls);
  call pgxls.add_row(xls);
  call pgxls.put_cell(xls, '10 sheets * 4 columns * 10K rows = 4M cells'::text);
  for v_sheet in 1..10 loop
    call pgxls.add_sheet(xls, array[10,10,10,50], array['Sheet','Row','Value','md5'], v_sheet::text);
    call pgxls.set_column_format(xls, 4, font_name=>pgxls.font_name$monospace());
    for v_row in 1..10000 loop          
      v_value := v_sheet*v_row;
      call pgxls.add_row(xls);
      call pgxls.put_cell_integer(xls, v_sheet);     
      call pgxls.put_cell_integer(xls, v_row);      
      call pgxls.put_cell_integer(xls, v_value);
      call pgxls.put_cell_text(xls, md5(v_value::text));     
    end loop;
  end loop;
  return query execute pgxls.get_file_parts_query(xls);
  call pgxls.clear_file_parts(xls);
end
$$;

-- Get large file by parts
select excel_large_file();

