-- postgres create wrapper for procedure pgxls.save_file with option security definer and access check,
-- grant privilege on wrapper    
create or replace procedure example.save_file_to_share_reports(inout xls pgxls.xls, filename varchar) security definer language plpgsql as $$  
begin 
  if filename like '%/%' then
    raise exception 'File name cannot contain path';   
  end if;
  call pgxls.save_file(xls, '/mnt/share/reports/'||filename);
end 
$$;
grant execute on procedure example.save_file_to_share_reports to developer_1;

-- developer_1 create report and save it using wrapper example.save_file_to_share_reports
create or replace procedure example.my_report() language plpgsql as $$
declare 
  xls pgxls.xls; 
begin
  xls := pgxls.create(array[10,10]);   
  call pgxls.put_cell_text(xls, 'My');      
  call pgxls.put_cell_text(xls, 'Report'); 
  call example.save_file_to_share_reports(xls, 'MyReport.xlsx');      
end
$$;
call example.my_report();
