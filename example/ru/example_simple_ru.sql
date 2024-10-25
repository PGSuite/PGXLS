-- 1. Создаем и получаем файл(bytea) по SQL-запросу
select pgxls.get_file_by_query('select * from pg_tables');

-- 2. Сохраняем Excel-файл на сервере по SQL-запросу
call pgxls.save_file_by_query('/tmp/top_relations_by_size.xlsx', 'select oid,relname,pg_relation_size(oid) from pg_class order by 3 desc limit 10');

-- 3. Создаем функцию, возвращающую файл(bytea)
create or replace function excel_top_relations_by_size() returns bytea language plpgsql as $$
declare 
  rec record;
  xls pgxls.xls; 
begin
  -- Создаем документ, в параметрах указываем ширины и заголовки столбцов      
  xls := pgxls.create(array[10,80,15], array['oid','Name','Size, bytes']);
  -- В цикле выбираем данные по запросу
  for rec in
    select oid,relname,pg_relation_size(oid) size from pg_class order by 3 desc limit 10    
  loop
    -- Добавляем строку
    call pgxls.add_row(xls);
    -- В ячейки устанавливаем данные из запроса
    call pgxls.put_cell(xls, rec.oid);      
    call pgxls.put_cell(xls, rec.relname);   
    call pgxls.put_cellvalue(xls, rec.size);
  end loop;  
  -- Возвращаем файл(bytea)
  return pgxls.get_file(xls);      
end
$$;
-- Получаем файл
select excel_top_relations_by_size();

