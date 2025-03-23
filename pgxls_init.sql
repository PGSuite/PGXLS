create schema if not exists pgxls;

create or replace function pgxls.pgxls_version() returns varchar language plpgsql as $$
begin
  return '24.4.14';
end; $$;

do $$ begin
  if to_regtype('pgxls.alignment_horizontal') is null then	
    create type pgxls.alignment_horizontal as enum ('left', 'center', 'right', 'justify', 'fill', 'distributed');
  end if;   
  if to_regtype('pgxls.alignment_vertical') is null then 
    create type pgxls.alignment_vertical as enum ('top', 'center', 'bottom', 'justify', 'distributed');
  end if;
  if to_regtype('pgxls.border_line') is null then
    create type pgxls.border_line as enum ('none', 'thin', 'thick', 'dashed', 'dotted', 'dashDot', 'dashDotDot', 'double');
  end if;
  if to_regtype('pgxls.page_paper_format') is null then
    create type pgxls.page_paper_format as enum ('A3','A4','A5');
  end if;
  if to_regtype('pgxls.page_orientation') is null then
    create type pgxls.page_orientation as enum ('portrait','landscape');
  end if;   
  --
  if to_regtype('pgxls._format') is null then
    create type pgxls._format as (
      code varchar(30)
    );  
  end if;
  --
  if to_regtype('pgxls._font') is null then
    create type pgxls._font as (
      name varchar(50),
      size int,
      bold boolean,
      italic boolean,
      underline boolean,
      strike boolean,
      color varchar(6)
    );  
  end if;
  --
  if to_regtype('pgxls._border') is null then
    create type pgxls._border as (
      left_ pgxls.border_line,
      top pgxls.border_line,
      right_ pgxls.border_line,
      bottom pgxls.border_line 
    );  
  end if;
  --
  if to_regtype('pgxls._fill') is null then
    create type pgxls._fill as (
      foreground_color varchar(6)
    );  
  end if;
  --
  if to_regtype('pgxls._style') is null then
    create type pgxls._style as (
      format int,
      font int,
      border int,
      fill int,
      alignment_horizontal pgxls.alignment_horizontal,
      alignment_indent int,      
      alignment_vertical pgxls.alignment_vertical,
      alignment_text_wrap boolean
    );  
  end if;
  --
  if to_regtype('pgxls._column') is null then
    create type pgxls._column as (
      styles int[],
      width int,      
      name varchar(4)
    );  
  end if;
  --
  if to_regtype('pgxls._cell') is null then
    create type pgxls._cell as (
      type char,
      style int,
      value text
    );  
  end if;
  --
  if to_regtype('pgxls._page') is null then
    create type pgxls._page as (
      header_alignment pgxls.alignment_horizontal,
      header_font_name varchar,
      header_font_size int,
      header_text text,
      margin_left numeric(10,3),
      margin_top numeric(10,3),
      margin_right numeric(10,3),
      margin_bottom numeric(10,3),
      rows_repeat_from int,
      rows_repeat_to int,
      paper_format pgxls.page_paper_format,
      paper_orientation pgxls.page_orientation
    );  
  end if;
  --
  if to_regtype('pgxls.xls') is null then
    create type pgxls.xls as (
      id int,      
      datetime timestamptz,
      --      
      formats pgxls._format[],
      fonts pgxls._font[],
      borders pgxls._border[],
      fills pgxls._fill[],
      styles pgxls._style[],
      page pgxls._page,
      --      
      columns pgxls._column[],
      columns_len int,
      column_default pgxls._column,
      column_current int,
      cells pgxls._cell[],
      cells_empty pgxls._cell[],
      row_height int,
      sheets_len int,
      sheet_file_name varchar(50),
      sheet_name varchar,
      rows_len int,
      cells_merge_len int,
      strings_len int, 
      --
      newline char,
      trace_len int,
      trace_ts timestamp
    );  
  end if;
end $$;

create or replace function pgxls.get_column_name(column_ int) returns varchar language plpgsql as $$
declare
  v_name varchar := '';
begin
  while column_!=0 loop
    v_name := chr(65+(column_-1)%26) || v_name;
    column_ := (column_-1)/26;
  end loop;  	
  return v_name;  
end
$$;

create or replace function pgxls.get_format_code_numeric(decimal_places int, thousands_separated boolean default true) returns varchar language plpgsql as $$
declare
  v_code varchar := '0';
begin
  if (decimal_places>0) then
    v_code := v_code||'.'||rpad('', decimal_places, '0');
  end if; 
  if thousands_separated then
    v_code := '#,##' || v_code;
  end if;
  return v_code;  
end
$$;

create or replace function pgxls.get_format_code_boolean(text_true varchar default 'True', text_false varchar default 'False', text_null varchar default '') returns varchar language plpgsql as $$
begin
  return '"'||text_true||'";"'||text_null||'";"'||text_false||'";';  
end
$$;

create or replace function pgxls.font_name$sans()       returns varchar language plpgsql as $$ begin return 'Arial';           end $$;
create or replace function pgxls.font_name$sans_serif() returns varchar language plpgsql as $$ begin return 'Times New Roman'; end $$;
create or replace function pgxls.font_name$monospace()  returns varchar language plpgsql as $$ begin return 'Courier New';     end $$;

create or replace function pgxls._color(red int, green int, blue int) returns varchar(6) language plpgsql as $$ begin return upper(lpad(to_hex(red & 255),2,'0')||lpad(to_hex(green & 255),2,'0')||lpad(to_hex(blue & 255),2,'0')); end $$;
create or replace function pgxls.color$light_red()   returns varchar(6) language plpgsql as $$ begin return pgxls._color(255,0,0);     end $$;
create or replace function pgxls.color$light_green() returns varchar(6) language plpgsql as $$ begin return pgxls._color(0,255,0);     end $$;
create or replace function pgxls.color$light_blue()  returns varchar(6) language plpgsql as $$ begin return pgxls._color(0,0,255);     end $$;
create or replace function pgxls.color$light_gray()  returns varchar(6) language plpgsql as $$ begin return pgxls._color(211,211,211); end $$;
create or replace function pgxls.color$dark_red()    returns varchar(6) language plpgsql as $$ begin return pgxls._color(128,0,0);     end $$;
create or replace function pgxls.color$dark_green()  returns varchar(6) language plpgsql as $$ begin return pgxls._color(0,128,0);     end $$;
create or replace function pgxls.color$dark_blue()   returns varchar(6) language plpgsql as $$ begin return pgxls._color(0,0,128);     end $$;
create or replace function pgxls.color$dark_gray()   returns varchar(6) language plpgsql as $$ begin return pgxls._color(169,169,169); end $$;

create or replace procedure pgxls._create_column_default(inout xls pgxls.xls) language plpgsql as $$
declare
  v_format pgxls._format;
  v_font pgxls._font;
  v_border pgxls._border;
  v_style pgxls._style; 
  v_column pgxls._column;
begin
  v_font.name      := pgxls.font_name$sans();
  v_font.size      := 10;
  v_font.bold      := false;
  v_font.italic    := false;
  v_font.underline := false;
  v_font.strike    := false;
  v_font.color     := 'auto'; 
  xls.fonts[0] := v_font;
  --
  v_border.left_ := 'none'; v_border.top := 'none'; v_border.right_ := 'none'; v_border.bottom := 'none';
  xls.borders[0] := v_border;
  --
  v_style.font                := 0;
  v_style.border              := 0;
  v_style.fill                := 0;
  v_style.alignment_indent    := 0;
  v_style.alignment_vertical  := 'center';
  v_style.alignment_text_wrap := false;
  --
  v_format.code := 'General';             xls.formats[100] := v_format; v_style.format := 100; v_style.alignment_horizontal := 'left';   xls.styles[0] := v_style; 
  v_format.code := '@';                   xls.formats[101] := v_format; v_style.format := 101; v_style.alignment_horizontal := 'left';   xls.styles[1] := v_style; -- text
  v_format.code := '0';                   xls.formats[102] := v_format; v_style.format := 102; v_style.alignment_horizontal := 'right';  xls.styles[2] := v_style; -- integer
  v_format.code := '0.00';                xls.formats[103] := v_format; v_style.format := 103; v_style.alignment_horizontal := 'right';  xls.styles[3] := v_style; -- numeric
  v_format.code := 'yyyy-mm-dd';          xls.formats[104] := v_format; v_style.format := 104; v_style.alignment_horizontal := 'center'; xls.styles[4] := v_style; -- date  
  v_format.code := 'hh:mm:ss';            xls.formats[105] := v_format; v_style.format := 105; v_style.alignment_horizontal := 'center'; xls.styles[5] := v_style; -- time  
  v_format.code := 'yyyy-mm-dd hh:mm:ss'; xls.formats[106] := v_format; v_style.format := 106; v_style.alignment_horizontal := 'center'; xls.styles[6] := v_style; -- timestamp
  v_format.code := '"True";"";"False";';  xls.formats[107] := v_format; v_style.format := 107; v_style.alignment_horizontal := 'center'; xls.styles[7] := v_style; -- boolean
  --
  for cell_type in 1..7 loop
    v_column.styles[cell_type] := cell_type; 
  end loop;
  -- 
  xls.column_default := v_column; 
end
$$;

create or replace function pgxls.create(columns_widths int[], columns_captions text[] default null, sheet_name varchar default null) returns pgxls.xls language plpgsql as $$
declare
  v_xls pgxls.xls;
  v_format pgxls._format;
  v_style pgxls._style;
begin
  v_xls.newline := chr(10); 
  if to_regtype('pgxls_temp_file') is null then	
    create temp sequence if not exists pgxls_id_seq cycle;	
    create temp table if not exists pgxls_temp_file(xls_id int, name varchar(32), part int, subpart int, body bytea not null);
  end if; 
  v_xls.id := nextval('pgxls_id_seq');  
  v_xls.datetime := now();
  call pgxls._build_file$rels(v_xls);
  v_xls.sheets_len := 0;
  v_xls.strings_len := 0;
  call pgxls._create_column_default(v_xls);
  call pgxls.add_sheet(v_xls, columns_widths, columns_captions, sheet_name);
  v_xls.trace_len := 0;
  v_xls.trace_ts := clock_timestamp();
  call pgxls._trace(v_xls, 'create', 'pgxls_version='||pgxls.pgxls_version()||', server_encoding='||current_setting('server_encoding')||', client_encoding='||current_setting('client_encoding'));
  return v_xls;	
end
$$;

create or replace procedure pgxls.add_sheet(inout xls pgxls.xls, columns_widths int[], columns_captions text[] default null, name varchar default null) language plpgsql as $$
declare
  v_column pgxls._column;
begin
  if xls.sheets_len>0 then 	
    call pgxls._build_file$xl_worksheets_sheet(xls);
  end if; 
  xls.columns_len := array_length(columns_widths,1);
  for c in 1..xls.columns_len loop
    v_column := xls.column_default;
    v_column.width := columns_widths[c];
    v_column.name := pgxls.get_column_name(c);
    xls.columns[c] := v_column;
  end loop;
  xls.cells_empty := array_fill(null::pgxls._cell, array[xls.columns_len]);
  xls.rows_len := 0;
  xls.cells_merge_len := 0;
  xls.sheets_len := xls.sheets_len+1;
  xls.sheet_file_name := 'xl/worksheets/sheet'||xls.sheets_len||'.xml';
  xls.sheet_name := coalesce(name, 'Sheet'||xls.sheets_len);
  if columns_captions is not null then 
    for c in 1..array_length(columns_captions,1) loop
      call pgxls.put_cell(xls, columns_captions[c], font_bold => true, alignment_horizontal => 'center');
    end loop;
    call pgxls._build_file$xl_worksheets_sheet_row(xls);   
  end if; 
  call pgxls.set_page_header(xls, 'PGSuite '||to_char(xls.datetime,'YYYY-MM-DD HH24:MI')||' &P#&N');
  call pgxls.set_page_margins(xls, 0.15, 0.2, 0.15, 0.4);
  call pgxls.set_page_rows_repeat(xls, case when columns_captions is not null then 1 else null end);
  call pgxls.set_page_paper(xls, 'A4', 'portrait');
end
$$;

create or replace procedure pgxls.add_sheet_by_query(inout xls pgxls.xls, query text, name varchar default null) language plpgsql as $$
declare
  v_rec_column record;
  v_columns_names varchar(256)[];
  v_columns_types regtype[];
  v_columns_widths int[];   
  v_columns_len int := 0;
  v_sql_block text;
begin
  call pgxls._trace(xls, 'add_sheet_by_query', 'started');
  if to_regtype('pgxls_query_data') is not null then
    drop table pgxls_query_data;
  end if;
  execute 'create temp table pgxls_query_data as select row_number() over()  _pgxls_rownum,t.* from ('||query||') t';
  v_sql_block :=
    'do $block$'||xls.newline||
    'declare '||xls.newline||
    '  xls pgxls.xls := (select var_xls from pgxls_query_block_var);'||xls.newline||
    '  rec pgxls_query_data%rowtype;'||xls.newline||
    'begin '||xls.newline||   
    '  for rec in select * from pgxls_query_data order by 1 loop'||xls.newline||
    '    call pgxls.add_row(xls);'||xls.newline;   
  for v_rec_column in  
    select attname,typcategory,typname::regtype
      from pg_attribute a
      join pg_type t on t.oid=a.atttypid
      where attrelid='pgxls_query_data'::regclass and attnum>1
      order by attnum
  loop
    v_columns_len := v_columns_len + 1;
    v_columns_names[v_columns_len] := v_rec_column.attname;
    v_columns_widths[v_columns_len] := 
      case when v_rec_column.typcategory='N' then 15
           when v_rec_column.typname in ('date','time','timetz','bool') then 12
           when v_rec_column.typname in ('timestamp','timestamptz') then 20
           else 35
      end;
    v_columns_widths[v_columns_len] := greatest(v_columns_widths[v_columns_len], length(v_rec_column.attname)*1.5);
    v_sql_block := v_sql_block||
      '    call pgxls.put_cell_'||
      case when pgxls._type_is_integer  (v_rec_column.typname) then 'integer   (xls, rec.'||quote_ident(v_rec_column.attname)||'::bigint'
           when pgxls._type_is_numeric  (v_rec_column.typname) then 'numeric   (xls, rec.'||quote_ident(v_rec_column.attname)||'::numeric'
           when pgxls._type_is_date     (v_rec_column.typname) then 'date      (xls, rec.'||quote_ident(v_rec_column.attname)||'::date'
           when pgxls._type_is_time     (v_rec_column.typname) then 'time      (xls, rec.'||quote_ident(v_rec_column.attname)||'::time'
           when pgxls._type_is_timestamp(v_rec_column.typname) then 'timestamp (xls, rec.'||quote_ident(v_rec_column.attname)||'::timestamp'
           when pgxls._type_is_boolean  (v_rec_column.typname) then 'boolean   (xls, rec.'||quote_ident(v_rec_column.attname)||'::boolean'
      else                                                          'text      (xls, rec.'||quote_ident(v_rec_column.attname)||'::text'
      end ||   
      ');'||xls.newline;   
  end loop;
  v_sql_block := v_sql_block||
    '  end loop;'||xls.newline||  
    '  update pgxls_query_block_var set var_xls=xls;'||xls.newline||
    'end'||xls.newline||
    '$block$'||xls.newline;
  call pgxls.add_sheet(xls, v_columns_widths, v_columns_names, name);
  if to_regtype('pgxls_query_block_var') is null then
    create temp table pgxls_query_block_var (var_xls pgxls.xls);
  else
    delete from pgxls_query_block_var;
  end if;  
  insert into pgxls_query_block_var values (xls);
  execute v_sql_block;
  xls := (select var_xls from pgxls_query_block_var);
  drop table pgxls_query_block_var;
  drop table pgxls_query_data;
  call pgxls._trace(xls, 'add_sheet_by_query', 'finished');
end
$$;

create or replace procedure pgxls.add_row(inout xls pgxls.xls, count int default 1) language plpgsql as $$
begin
  for i in 1..count loop	
    if xls.cells is not null then
      call pgxls._build_file$xl_worksheets_sheet_row(xls);
    end if;
    xls.cells := xls.cells_empty;  
    xls.rows_len := xls.rows_len+1;
  end loop; 
  xls.column_current := 0;
end
$$;

create or replace procedure pgxls.add_row_texts(inout xls pgxls.xls, texts text[], font_bold boolean default null, fill_foreground_color varchar(6) default null, alignment_horizontal pgxls.alignment_horizontal default null) language plpgsql as $$
declare
  v_text text;
begin
  call pgxls.add_row(xls);	
  foreach v_text in array texts loop
    call pgxls.put_cell_text(xls, v_text); 
  end loop;
  call pgxls.format_row(xls, font_bold=>font_bold, fill_foreground_color=>fill_foreground_color, alignment_horizontal=>alignment_horizontal);
end
$$;

create or replace procedure pgxls.next_column(inout xls pgxls.xls) language plpgsql as $$
begin
  xls.column_current := xls.column_current+1;
end
$$;

create or replace procedure pgxls._next_column_default(inout xls pgxls.xls, column_ int) language plpgsql as $$
begin
  if xls.cells is null then
    call pgxls.add_row(xls);
  end if;
  if column_ is not null then
    xls.column_current := column_; 
  else 
    call pgxls.next_column(xls);
  end if; 
end
$$;

create or replace procedure pgxls.set_column_current(inout xls pgxls.xls, column_ int) language plpgsql as $$
begin
  xls.column_current := column_;
end
$$;

create or replace procedure pgxls._add_style(inout xls pgxls.xls, inout style int, format int default null, font int default null, border int default null, fill int default null, alignment_horizontal pgxls.alignment_horizontal default null, alignment_indent int default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null) language plpgsql as $$
declare
  v_styles_len int := array_length(xls.styles,1);
  v_style pgxls._style := xls.styles[style]; 
begin
  if format               is not null then v_style.format               := format;               end if;
  if font                 is not null then v_style.font                 := font;                 end if;
  if border               is not null then v_style.border               := border;               end if;
  if fill                 is not null then v_style.fill                 := fill;                 end if;
  if alignment_horizontal is not null then v_style.alignment_horizontal := alignment_horizontal; end if;
  if alignment_indent     is not null then v_style.alignment_indent     := alignment_indent;     end if;
  if alignment_vertical   is not null then v_style.alignment_vertical   := alignment_vertical;   end if;
  if alignment_text_wrap  is not null then v_style.alignment_text_wrap  := alignment_text_wrap;  end if; 
  for s in 0..v_styles_len-1 loop
    if 
      xls.styles[s].format               = v_style.format               and
      xls.styles[s].font                 = v_style.font                 and 
      xls.styles[s].border               = v_style.border               and      
      xls.styles[s].fill                 = v_style.fill                 and
      xls.styles[s].alignment_horizontal = v_style.alignment_horizontal and  
      xls.styles[s].alignment_indent     = v_style.alignment_indent     and      
      xls.styles[s].alignment_vertical   = v_style.alignment_vertical   and
      xls.styles[s].alignment_text_wrap  = v_style.alignment_text_wrap      
    then
      style := s;
      return;
    end if;
  end loop;
  style := v_styles_len;
  xls.styles[style] = v_style;
end
$$;

create or replace procedure pgxls._add_format_code(inout xls pgxls.xls, inout format int, code varchar) language plpgsql as $$
declare
  v_format pgxls._format;
begin
  for f in array_lower(xls.formats,1)..array_upper(xls.formats,1) loop	
    if xls.formats[f].code=code then
      format := f;
      return;
    end if;
  end loop;
  format := array_upper(xls.formats,1)+1;
  v_format.code := code;
  xls.formats[format] := v_format;
end
$$;

create or replace procedure pgxls._add_font(inout xls pgxls.xls, inout font int, name varchar, size int, bold boolean, italic boolean, underline boolean, strike boolean, color varchar(6)) language plpgsql as $$
declare
  v_fonts_len int := array_length(xls.fonts,1);
  v_font pgxls._font := xls.fonts[font]; 
begin
  if name      is not null then v_font.name      := name;     end if;
  if size      is not null then v_font.size      := size;     end if;
  if bold      is not null then v_font.bold      := bold;     end if;
  if italic    is not null then v_font.italic    :=italic;    end if;
  if underline is not null then v_font.underline :=underline; end if;
  if strike    is not null then v_font.strike    := strike;   end if;
  if color     is not null then v_font.color     := color;    end if; 
  for f in 0..v_fonts_len-1 loop
    if 
      xls.fonts[f].name      = v_font.name      and
      xls.fonts[f].size      = v_font.size      and
      xls.fonts[f].bold      = v_font.bold      and
      xls.fonts[f].italic    = v_font.italic    and
      xls.fonts[f].underline = v_font.underline and
      xls.fonts[f].strike    = v_font.strike    and
      xls.fonts[f].color     = v_font.color 
    then
      font := f;
      return;
    end if;
  end loop;
  font := v_fonts_len;
  xls.fonts[font] = v_font;
end
$$;

create or replace procedure pgxls._add_border(inout xls pgxls.xls, inout border int, around pgxls.border_line, left_ pgxls.border_line, top pgxls.border_line, right_ pgxls.border_line, bottom pgxls.border_line) language plpgsql as $$
declare
  v_borders_len int := array_length(xls.borders,1);
  v_border pgxls._border := xls.borders[border]; 
begin
  if around is not null then 
    v_border.left_ := around; v_border.top := around; v_border.right_ := around; v_border.bottom := around;
  else
    if left_  is not null then v_border.left_  := left_;  end if;
    if top    is not null then v_border.top    := top;    end if;
    if right_ is not null then v_border.right_ := right_; end if;
    if bottom is not null then v_border.bottom := bottom; end if;
  end if; 
  for b in 0..v_borders_len-1 loop	
    if xls.borders[b].left_ = v_border.left_ and xls.borders[b].top = v_border.top and xls.borders[b].right_ = v_border.right_ and xls.borders[b].bottom = v_border.bottom then
      border := b;
      return;
    end if;
  end loop;
  border := v_borders_len;
  xls.borders[v_borders_len] = v_border;
end
$$;

create or replace procedure pgxls._add_fill(inout xls pgxls.xls, inout fill int, foreground_color varchar(6)) language plpgsql as $$
declare
  v_fills_upper int := coalesce(array_upper(xls.fills,1),1);
  v_fill pgxls._fill;
begin
  if foreground_color is null or length(foreground_color)<6 then
    fill := 0;
    foreground_color := 'none';
  end if;
  for f in 2..v_fills_upper loop	
    if xls.fills[f].foreground_color=foreground_color then
      fill := f;
      return;
    end if;
  end loop;
  fill := v_fills_upper+1;
  v_fill.foreground_color := foreground_color;
  xls.fills[fill] := v_fill;
end
$$;

do $body$
declare 
  v_type record;  
  v_newline char := chr(10);  
  v_func_name text;
  v_func_params_def text :=
    'font_name varchar default null, font_size int default null, font_bold boolean default null, font_italic boolean default null, font_underline boolean default null, font_strike boolean default null, font_color varchar(6) default null, '
    'border_around pgxls.border_line default null, border_left pgxls.border_line default null, border_top pgxls.border_line default null, border_right pgxls.border_line default null, border_bottom pgxls.border_line default null, '
    'fill_foreground_color varchar(6) default null, ' 
    'alignment_horizontal pgxls.alignment_horizontal default null, alignment_indent int default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null';  
  v_func_params_call text :=
    'font_name, font_size, font_bold, font_italic, font_underline, font_strike, font_color, '
    'border_around, border_left, border_top, border_right, border_bottom, '
    'fill_foreground_color, '
    'alignment_horizontal, alignment_indent, alignment_vertical, alignment_text_wrap';
  v_funcs_call text := '';   
begin
  for v_type in (select * from unnest(array['text','integer','numeric','date','time','timestamp','boolean']) with ordinality as t(name, position)) loop
      v_func_name := 'pgxls.set_column_format_'||v_type.name;
    execute  
      'create or replace procedure '||v_func_name||'(inout xls pgxls.xls, column_ int, format_code varchar default null, '||v_func_params_def||') language plpgsql as $$'||v_newline||
      'declare'||v_newline||
      '  v_column pgxls._column := xls.columns[column_];'||v_newline||
      '  v_style int := v_column.styles['||v_type.position||'];'||v_newline||
      '  v_format int;'||v_newline||
      '  v_font int;'||v_newline||      
      '  v_border int;'||v_newline||
      '  v_fill int;'||v_newline||
      'begin'||v_newline||
      '  if format_code is not null then'||v_newline||
      '    v_format := null;'||v_newline||
      '    call pgxls._add_format_code(xls, v_format, format_code);'||v_newline||
      '  end if;'||v_newline||
      '  if font_name is not null or font_size is not null or font_bold is not null or font_italic is not null or font_underline is not null or font_strike is not null or font_color is not null then'||v_newline||
      '    v_font := xls.styles[v_style].font;'||v_newline||
      '    call pgxls._add_font(xls, v_font, font_name, font_size, font_bold, font_italic, font_underline, font_strike, font_color);'||v_newline||
      '  end if;'||v_newline||
      '  if border_around is not null or border_left is not null or border_top is not null or border_right is not null or border_bottom is not null then'||v_newline||
      '    v_border := xls.styles[v_style].border;'||v_newline||
      '    call pgxls._add_border(xls, v_border, border_around, border_left, border_top, border_right, border_bottom);'||v_newline||
      '  end if;'||v_newline||      
      '  if fill_foreground_color is not null then'||v_newline||
      '    v_fill := xls.styles[v_style].fill;'||v_newline||
      '    call pgxls._add_fill(xls, v_fill, fill_foreground_color);'||v_newline||
      '  end if;'||v_newline|| 
      '  call pgxls._add_style(xls, v_style, v_format, v_font, v_border, v_fill, alignment_horizontal, alignment_indent, alignment_vertical, alignment_text_wrap);'||v_newline||
      '  v_column.styles['||v_type.position||'] := v_style;'||v_newline||
      '  xls.columns[column_] := v_column;'||v_newline||
      'end'||v_newline||
    '$$;';
    v_funcs_call := v_funcs_call || '  call '||v_func_name||'(xls, column_, null, '||v_func_params_call||');'||v_newline;
  end loop; 
  execute
    'create or replace procedure pgxls.set_column_format(inout xls pgxls.xls, column_ int, '||v_func_params_def||') language plpgsql as $$'||v_newline||
    'begin'||v_newline||
    v_funcs_call||
    'end'||v_newline||
    '$$;'; 
end
$body$;

create or replace procedure pgxls._set_cell_style(inout xls pgxls.xls, format int default null, font int default null, border int default null, fill int default null, alignment_horizontal pgxls.alignment_horizontal default null, alignment_indent int default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null, column_ int default null) language plpgsql as $$
declare 
  v_column int := coalesce(column_, xls.column_current); 
  v_style int;
  v_cell pgxls._cell;
begin
  if v_column<1 or v_column>xls.columns_len then 	
  	raise exception 'Column % out of range [1,%]', v_column, xls.columns_len;
  end if;
  v_cell := xls.cells[v_column];
  if v_cell is null then
    v_cell.type := 's';
    v_cell.style := xls.columns[v_column].styles[1];
    v_cell.value := '';
  end if;
  v_style := v_cell.style;
  call pgxls._add_style(xls, v_style, format, font, border, fill, alignment_horizontal, alignment_indent, alignment_vertical, alignment_text_wrap);
  v_cell.style := v_style;
  xls.cells[v_column] := v_cell;
end
$$;

create or replace function pgxls._type_is_integer  (type regtype) returns boolean language plpgsql as $$ begin return type in ('integer'::regtype,'bigint'::regtype,'smallint'::regtype,'oid'::regtype); end $$;
create or replace function pgxls._type_is_numeric  (type regtype) returns boolean language plpgsql as $$ begin return type in ('numeric'::regtype,'real'::regtype); end $$;
create or replace function pgxls._type_is_date     (type regtype) returns boolean language plpgsql as $$ begin return type in ('date'::regtype); end $$;
create or replace function pgxls._type_is_time     (type regtype) returns boolean language plpgsql as $$ begin return type in ('time'::regtype,'timetz'::regtype); end $$;
create or replace function pgxls._type_is_timestamp(type regtype) returns boolean language plpgsql as $$ begin return type in ('timestamp'::regtype,'timestamptz'::regtype); end $$;
create or replace function pgxls._type_is_boolean  (type regtype) returns boolean language plpgsql as $$ begin return type in ('boolean'::regtype); end $$;

create or replace procedure pgxls.put_cell_text(inout xls pgxls.xls, value text, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);
  v_cell.type := 's';
  v_cell.style := xls.columns[xls.column_current].styles[1];
  v_cell.value := coalesce(value,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_integer(inout xls pgxls.xls, value bigint, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[2];
  v_cell.value := coalesce(value::text,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_numeric(inout xls pgxls.xls, value numeric, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[3];
  v_cell.value := coalesce(value::text,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_date(inout xls pgxls.xls, value date, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[4];
  v_cell.value := coalesce((extract(epoch from value)/24/60/60+25569)::text,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_time(inout xls pgxls.xls, value time, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[5];
  v_cell.value := coalesce((extract(epoch from value)/24.0/60/60)::text,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_timestamp(inout xls pgxls.xls, value timestamp, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[6];
  v_cell.value := coalesce((extract(epoch from value)/24.0/60/60+25569)::text,'');
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.put_cell_boolean(inout xls pgxls.xls, value boolean, column_ int default null) language plpgsql as $$
declare
  v_cell pgxls._cell;
begin
  call pgxls._next_column_default(xls, column_);	
  v_cell.type := 'n';  
  v_cell.style := xls.columns[xls.column_current].styles[7];
  v_cell.value := coalesce(value::int,-1);
  xls.cells[xls.column_current] := v_cell;
end
$$;

create or replace procedure pgxls.format_cell(
  inout xls pgxls.xls, 
  column_ int default null,
  format_code varchar default null,
  font_name varchar default null, font_size int default null, font_bold boolean default null, font_italic boolean default null, font_underline boolean default null, font_strike boolean default null, font_color varchar(6) default null,
  border_around pgxls.border_line default null, border_left pgxls.border_line default null, border_top pgxls.border_line default null, border_right pgxls.border_line default null, border_bottom pgxls.border_line default null,
  fill_foreground_color varchar(6) default null, 
  alignment_horizontal pgxls.alignment_horizontal default null, alignment_indent int default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null  
) language plpgsql as $$
declare
  v_column int := coalesce(column_, xls.column_current);
  v_format int;
  v_font int;
  v_border int;
  v_fill int; 
begin
  if format_code is not null then    
    call pgxls._add_format_code(xls, v_format, format_code);
    call pgxls._set_cell_style(xls, format => v_format, column_=>v_column);   
  end if;
  if font_name is not null or font_size is not null or font_bold is not null or font_italic is not null or font_underline is not null or font_strike is not null or font_color is not null then
    v_font := xls.styles[xls.cells[v_column].style].font;
    call pgxls._add_font(xls, v_font, font_name, font_size, font_bold, font_italic, font_underline, font_strike, font_color);
    call pgxls._set_cell_style(xls, font => v_font, column_=>v_column); 
  end if;
  if border_around is not null or border_left is not null or border_top is not null or border_right is not null or border_bottom is not null then
    v_border := xls.styles[xls.cells[v_column].style].border;
    call pgxls._add_border(xls, v_border, border_around, border_left, border_top, border_right, border_bottom);
    call pgxls._set_cell_style(xls, border => v_border, column_=>v_column);
  end if;
  if fill_foreground_color is not null then
    call pgxls._add_fill(xls, v_fill, fill_foreground_color);
    call pgxls._set_cell_style(xls, fill=>v_fill, column_=>v_column); 
  end if; 
  if alignment_horizontal is not null or alignment_indent is not null or alignment_vertical is not null or alignment_text_wrap is not null then
    call pgxls._set_cell_style(xls, alignment_horizontal=>alignment_horizontal, alignment_indent=>alignment_indent, alignment_vertical=>alignment_vertical, alignment_text_wrap=>alignment_text_wrap, column_=>v_column); 	
  end if;
end
$$;

create or replace procedure pgxls.put_cell(
  inout xls pgxls.xls, 
  value anyelement, 
  column_ int default null,
  format_code varchar default null,
  font_name varchar default null, font_size int default null, font_bold boolean default null, font_italic boolean default null, font_underline boolean default null, font_strike boolean default null, font_color varchar(6) default null,
  border_around pgxls.border_line default null, border_left pgxls.border_line default null, border_top pgxls.border_line default null, border_right pgxls.border_line default null, border_bottom pgxls.border_line default null,
  fill_foreground_color varchar(6) default null, 
  alignment_horizontal pgxls.alignment_horizontal default null, alignment_indent int default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null  
) language plpgsql as $$
declare 
  v_column int; 
  v_cell pgxls._cell;
  v_value_type regtype := pg_typeof(value);
begin
  call pgxls._next_column_default(xls, column_);
  if xls.column_current<1 or xls.column_current>xls.columns_len then 	
  	raise exception 'Column % out of range [1,%]', xls.column_current, xls.columns_len;
  end if;
  if     pgxls._type_is_integer  (v_value_type) then call pgxls.put_cell_integer   (xls, value::bigint,    xls.column_current);
  elseif pgxls._type_is_numeric  (v_value_type) then call pgxls.put_cell_numeric   (xls, value::numeric,   xls.column_current);
  elseif pgxls._type_is_date     (v_value_type) then call pgxls.put_cell_date      (xls, value::date,      xls.column_current);
  elseif pgxls._type_is_time     (v_value_type) then call pgxls.put_cell_time      (xls, value::time,      xls.column_current);
  elseif pgxls._type_is_timestamp(v_value_type) then call pgxls.put_cell_timestamp (xls, value::timestamp, xls.column_current);
  elseif pgxls._type_is_boolean  (v_value_type) then call pgxls.put_cell_boolean   (xls, value::boolean,   xls.column_current);  
  else                                               call pgxls.put_cell_text      (xls, value::text,      xls.column_current);  
  end if;
  call pgxls.format_cell(xls, xls.column_current, format_code, font_name, font_size, font_bold, font_italic, font_underline, font_strike, font_color, border_around, border_left, border_top, border_right, border_bottom, fill_foreground_color, alignment_horizontal, alignment_indent, alignment_vertical, alignment_text_wrap);
end
$$;

create or replace procedure pgxls.format_row(
  inout xls pgxls.xls, 
  font_name varchar default null, font_size int default null, font_bold boolean default null, font_color varchar(6) default null,
  border pgxls.border_line default null, 
  fill_foreground_color varchar(6) default null, 
  alignment_horizontal pgxls.alignment_horizontal default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null  
) language plpgsql as $$
begin
  if xls.cells is null then 	
  	raise exception 'Row not added, call pgxls.add_row first';
  end if;
  for column_ in 1..xls.columns_len loop
    call pgxls.format_cell(xls, column_, null, font_name, font_size, font_bold, null, null, null, font_color, border, null, null, null, null, fill_foreground_color, alignment_horizontal, null, alignment_vertical, alignment_text_wrap);
  end loop;
end
$$;

create or replace procedure pgxls.set_all_format(
  inout xls pgxls.xls,
  font_name varchar default null, font_size int default null, font_bold boolean default null, font_color varchar(6) default null,
  border pgxls.border_line default null, 
  fill_foreground_color varchar(6) default null, 
  alignment_horizontal pgxls.alignment_horizontal default null, alignment_vertical pgxls.alignment_vertical default null, alignment_text_wrap boolean default null  
) language plpgsql as $$
begin
  -- if xls.cells is not null then
  --   call pgxls.format_row(xls, font_name, font_size, font_bold, font_color, border, fill_foreground_color, alignment_horizontal, alignment_vertical, alignment_text_wrap);
  -- end if;
  for column_ in 1..xls.columns_len loop
    call pgxls.set_column_format(xls, column_, font_name, font_size, font_bold, null, null, null, font_color, border, null, null, null, null, fill_foreground_color, alignment_horizontal, null, alignment_vertical, alignment_text_wrap);  
  end loop;
end
$$;

create or replace procedure pgxls.merge_cells(inout xls pgxls.xls, column_count int default 1, row_count int default 1, column_ int default null) language plpgsql as $$
begin
  call pgxls._build_file$xl_worksheets_sheet_merge_cells(xls, column_count, row_count, coalesce(column_,xls.column_current));
end
$$;

create or replace procedure pgxls.set_row_height(inout xls pgxls.xls, height int) language plpgsql as $$
begin
  if xls.cells is null then
    call pgxls.add_row(xls);
  end if;
  xls.row_height := height;
end
$$;

create or replace function pgxls.get_cell_height(line_count int default 1, font_size int default 10) returns int language plpgsql as $$
begin
  return 1+font_size*1.3*line_count;
end
$$;

create or replace procedure pgxls.set_page_header(inout xls pgxls.xls, header text, alignment pgxls.alignment_horizontal default 'right', font_name varchar default pgxls.font_name$sans(), font_size int default 6) language plpgsql as $$
declare
  v_page pgxls._page := xls.page;
begin
  v_page.header_alignment := alignment;
  v_page.header_font_name := font_name;
  v_page.header_font_size := font_size;
  v_page.header_text      := header;
  xls.page := v_page;
end
$$;

create or replace procedure pgxls.set_page_margins(inout xls pgxls.xls, left_ numeric default null, top numeric default null, right_ numeric default null, bottom numeric default null) language plpgsql as $$
declare
  v_page pgxls._page := xls.page;
begin
  if left_  is not null then v_page.margin_left   := left_;  end if;
  if top    is not null then v_page.margin_top    := top;    end if;
  if right_ is not null then v_page.margin_right  := right_; end if;
  if bottom is not null then v_page.margin_bottom := bottom; end if;
  xls.page := v_page;
end
$$;

create or replace procedure pgxls.set_page_rows_repeat(inout xls pgxls.xls, row_from int, row_to int default null) language plpgsql as $$
declare
  v_page pgxls._page := xls.page;
begin
  v_page.rows_repeat_from := row_from;
  v_page.rows_repeat_to   := coalesce(row_to, row_from);
  xls.page := v_page;
end
$$;

create or replace procedure pgxls.set_page_paper(inout xls pgxls.xls, format pgxls.page_paper_format default null, orientation pgxls.page_orientation default null) language plpgsql as $$
declare
  v_page pgxls._page := xls.page;
begin
  if format      is not null then v_page.paper_format      := format;      end if;
  if orientation is not null then v_page.paper_orientation := orientation; end if;
  xls.page := v_page;
end
$$;


create or replace procedure pgxls._build_file(inout xls pgxls.xls) language plpgsql as $$
begin
  if not exists (select from pgxls_temp_file where xls_id = xls.id and name='_rels/.rels') then
    raise exception 'Temporary table cleared (xls.id = %)', xls.id;
  end if;
  if exists (select from pgxls_temp_file where xls_id = xls.id and name='docProps/app.xml') then
    return;
  end if;
  call pgxls._trace(xls, '_build_file', 'started');
  call pgxls._build_file$docprops_app(xls);
  call pgxls._build_file$docprops_core(xls); 
  call pgxls._build_file$xl_rels_workbook(xls);
  call pgxls._build_file$xl_styles(xls);
  call pgxls._build_file$xl_worksheets_sheet(xls);
  call pgxls._build_file$xl_shared_strings(xls); 
  call pgxls._build_file$xl_workbook(xls);
  call pgxls._build_file$content_types(xls);
  call pgxls._zip_build(xls);
end
$$;

create or replace function pgxls.get_file(xls pgxls.xls) returns bytea language plpgsql as $$
declare
  v_file bytea;
begin
  call pgxls._build_file(xls);	
  v_file := (select string_agg(body, null order by name collate "C", part, subpart) from pgxls_temp_file where xls_id=xls.id);
  call pgxls.clear_file_parts(xls); 
  return v_file;	
end
$$;

create or replace function pgxls.get_file_parts_query(xls pgxls.xls) returns varchar language plpgsql as $$
begin	
  call pgxls._build_file(xls);
  return 'select body from pgxls_temp_file where xls_id='||xls.id||' order by name collate "C", part, subpart';
end
$$;

create or replace procedure pgxls.clear_file_parts(inout xls pgxls.xls) language plpgsql as $$
begin
  delete from pgxls_temp_file where xls_id=xls.id;  
end
$$;

create or replace function pgxls.get_file_by_query(query text) returns bytea language plpgsql as $$
declare
  xls pgxls.xls;
begin
  xls := pgxls.create(array[10]);
  xls.sheets_len := 0;
  call pgxls.add_sheet_by_query(xls, query);
  return pgxls.get_file(xls);	
end
$$;

create or replace procedure pgxls.save_file(inout xls pgxls.xls, filepath varchar) language plpgsql as $$  
declare
  v_lo oid;
begin 
  v_lo := lo_from_bytea(0, pgxls.get_file(xls));
  perform lo_export(v_lo, filepath);
  perform lo_unlink(v_lo);
end 
$$;

create or replace procedure pgxls.save_file_by_query(filepath varchar, query text) language plpgsql as $$
declare
  v_lo oid;
begin 
  v_lo := lo_from_bytea(0, pgxls.get_file_by_query(query));
  perform lo_export(v_lo, filepath);
  perform lo_unlink(v_lo);
end 
$$;

create or replace procedure pgxls._add_file_subpart(xls_id int, name varchar, part int, subpart int, body text) language plpgsql as $$
begin
  insert into pgxls_temp_file(xls_id, name, part, subpart, body) values (xls_id, name, part, subpart, pgxls._zip_utf8_bytea(body));
end;
$$; 

create or replace procedure pgxls._trace(inout xls pgxls.xls, proc_name name, log text) language plpgsql as $$
declare
  v_trace_ts_new timestamp := clock_timestamp();
begin
  xls.trace_len := xls.trace_len+1;
  call pgxls._add_file_subpart(xls.id, 'docProps/core.xml', 2, xls.trace_len,
    '    '||lpad(extract(epoch from v_trace_ts_new-xls.trace_ts)::numeric(10,3)::text,6,'0')||' '||rpad(proc_name, 20)||' '||log||xls.newline
  );
  xls.trace_ts := v_trace_ts_new;
end;
$$; 

create or replace procedure pgxls._build_file$rels(inout xls pgxls.xls) language plpgsql as $$
begin
  call pgxls._add_file_subpart(xls.id, '_rels/.rels', 1, 1, 	
    '<?xml version="1.0" encoding="UTF-8"?>'||xls.newline||
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'||xls.newline||
    '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'||xls.newline||
    '  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>'||xls.newline||
    '  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>'||xls.newline||
    '</Relationships>'  
  );
end
$$;

create or replace procedure pgxls._build_file$docprops_app(inout xls pgxls.xls) language plpgsql as $$
begin
  call pgxls._add_file_subpart(xls.id, 'docProps/app.xml', 1, 1, 	
    '<?xml version="1.0" encoding="UTF-8"?>'||xls.newline||
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'||xls.newline||
    '  <Template></Template>'||xls.newline||
    '  <TotalTime>0</TotalTime>'||xls.newline||
    '  <Application>PGXLS 0.0</Application>'||xls.newline||
    '</Properties>'
  ); 
end
$$;

create or replace procedure pgxls._build_file$docprops_core(inout xls pgxls.xls) language plpgsql as $$
declare
  v_file_datetime varchar(20) := to_char(now(), 'yyyy-mm-ddThh24:mi:ssZ'); 
begin
  call pgxls._add_file_subpart(xls.id, 'docProps/core.xml', 1, 1, 	
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||xls.newline||
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'||xls.newline||
    '  <dcterms:created xsi:type="dcterms:W3CDTF">'||v_file_datetime||'</dcterms:created>'||xls.newline||
    '  <dc:creator></dc:creator>'||xls.newline||
    '  <dc:description></dc:description>'||xls.newline||
    '  <dc:language>en</dc:language>'||xls.newline||
    '  <cp:lastModifiedBy></cp:lastModifiedBy>'||xls.newline||
    '  <dcterms:modified xsi:type="dcterms:W3CDTF">'||v_file_datetime||'</dcterms:modified>'||xls.newline||
    '  <cp:revision>1</cp:revision>'||xls.newline||
    '  <dc:subject></dc:subject>'||xls.newline||
    '  <dc:title></dc:title>'||xls.newline||
    '  <!-- trace'||xls.newline    
  );
  call pgxls._add_file_subpart(xls.id, 'docProps/core.xml', 9, 1, 	
    '  -->'||xls.newline||
    '</cp:coreProperties>'
  );
end
$$;

create or replace procedure pgxls._build_file$xl_rels_workbook(inout xls pgxls.xls) language plpgsql as $$
declare
  v_body text; 
begin
  v_body :=
    '<?xml version="1.0" encoding="UTF-8"?>'||xls.newline||
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'||xls.newline||
    '  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'||xls.newline;
  for v_sheet in 1..xls.sheets_len loop   
    v_body := v_body ||
      '  <Relationship Id="rId'||(v_sheet+1)||'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'||v_sheet||'.xml"/>'||xls.newline;
  end loop;
  v_body := v_body || 
      '  <Relationship Id="rId'||(xls.sheets_len+2)||'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />'||xls.newline|| 
      '</Relationships>';
  call pgxls._add_file_subpart(xls.id, 'xl/_rels/workbook.xml.rels', 1, 1, v_body);   
end
$$;

create or replace procedure pgxls._build_file$xl_worksheets_sheet_row(inout xls pgxls.xls) language plpgsql as $$
declare
  v_body text := '';
  v_cell pgxls._cell;
  v_style pgxls._style;
  v_font_size int;
  v_line text; 
  v_line_count int; 
  v_row_height int := 0;
  v_cell_height int;
begin
  for v_column in 1..xls.columns_len loop
    v_cell := xls.cells[v_column];
    continue when v_cell is null;
    if xls.row_height is null then
      v_style := xls.styles[xls.cells[v_column].style];
      v_font_size := xls.fonts[v_style.font].size;
      if v_style.alignment_text_wrap then
       v_cell.value := translate(v_cell.value,chr(13),'');
       v_line_count := 0; 
       foreach v_line in array string_to_array(v_cell.value,chr(10)) loop
         v_line_count := v_line_count + ceil(length(v_line)/(xls.columns[v_column].width*10.0/v_font_size-v_style.alignment_indent-3)+0.1);
       end loop;
      else
        v_line_count := 1;
      end if;
      v_cell_height := pgxls.get_cell_height(v_line_count, v_font_size);
      if v_row_height < v_cell_height then
        v_row_height := v_cell_height;
      end if;
    end if; 
    if v_cell.type='s' and v_cell.value!=''  then   
      call pgxls._add_file_subpart(xls.id, 'xl/sharedStrings.xml', 2, xls.strings_len,
        '  <si>'||xmlelement(name "t", xmlattributes('preserve' as "xml:space"), v_cell.value)||'</si>'||xls.newline	
      );    
      v_cell.value := xls.strings_len;
      xls.strings_len := xls.strings_len+1;
    end if;     
    v_body := v_body || '      <c r="'||xls.columns[v_column].name||xls.rows_len||'" s="'||v_cell.style||'" t="'||v_cell.type||'"><v>'||v_cell.value||'</v></c>'||xls.newline;
  end loop;
  call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 2, xls.rows_len,
    '    <row r="'||xls.rows_len||'" customFormat="false" customHeight="'||(xls.row_height is not null or v_row_height>0)||'" ht="'||coalesce(xls.row_height,v_row_height)||'" hidden="false" outlineLevel="0" collapsed="false">'||xls.newline||  
    v_body||
    '    </row>'||xls.newline
  );
  xls.cells := null;
  xls.row_height := null;
end
$$;

create or replace procedure pgxls._build_file$xl_worksheets_sheet_merge_cells(inout xls pgxls.xls, column_count int, row_count int, column_ int) language plpgsql as $$
begin
  xls.cells_merge_len := xls.cells_merge_len+1; 	
  call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 4, xls.cells_merge_len, 
    '  <mergeCell ref="'||pgxls.get_column_name(column_)||xls.rows_len||':'||pgxls.get_column_name(column_+column_count-1)||(xls.rows_len+row_count-1)||'"/>'||xls.newline
  );
end
$$;

create or replace procedure pgxls._build_file$xl_worksheets_sheet(inout xls pgxls.xls) language plpgsql as $$
declare
  v_body text;
  v_page pgxls._page := xls.page;
begin
  if xls.rows_len = 0 then
    call pgxls.add_row(xls);
  end if;
  if xls.cells is not null then
    call pgxls._build_file$xl_worksheets_sheet_row(xls);          
  end if; 
  call pgxls._add_file_subpart(xls.id, 'xl/workbook.xml', 2, xls.sheets_len, 
      '    '||xmlelement(name "sheet", xmlattributes(xls.sheet_name as "name", xls.sheets_len as "sheetId", 'visible' as "state", 'rId'||(xls.sheets_len+1) as "r:id"))||xls.newline
  );
  if v_page.rows_repeat_from is not null then
    call pgxls._add_file_subpart(xls.id, 'xl/workbook.xml', 4, xls.sheets_len,
       '    '||replace(xmlelement(name "definedName", xmlattributes('_xlnm.Print_Titles' as "name", (xls.sheets_len-1) as "localSheetId"), quote_literal(xls.sheet_name)||'!$'||v_page.rows_repeat_from||':$'||v_page.rows_repeat_to)::text,'''','&apos;')||xls.newline
    );
  end if;
  v_body := 
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||xls.newline||    
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'||xls.newline||
    '  <sheetPr filterMode="false"><pageSetUpPr fitToPage="false"/></sheetPr>'||xls.newline||
    '  <dimension ref="A1:'||xls.columns[xls.columns_len].name||xls.rows_len||'"/>'||xls.newline||
    '  <sheetViews>'||xls.newline||
    '    <sheetView showFormulas="false" showGridLines="true" showRowColHeaders="true" showZeros="true" rightToLeft="false" tabSelected="'||(xls.sheets_len=1)||'" showOutlineSymbols="true" defaultGridColor="true" view="normal" topLeftCell="A1" colorId="64" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100" workbookViewId="0">'||xls.newline||
    '      <selection pane="topLeft" activeCell="A1" activeCellId="0" sqref="A1"/>'||xls.newline||
    '    </sheetView>'||xls.newline||
    '  </sheetViews>'||xls.newline||
    '  <sheetFormatPr defaultColWidth="12" defaultRowHeight="12" zeroHeight="false" outlineLevelRow="0" outlineLevelCol="0"></sheetFormatPr>'||xls.newline||      
    '  <cols>'||xls.newline;
  for v_column in 1..xls.columns_len loop
    v_body := v_body || '    <col collapsed="false" customWidth="true" hidden="false" outlineLevel="0" max="'||v_column||'" min="'||v_column||'" style="0" width="'||xls.columns[v_column].width||'" />'||xls.newline;
  end loop;         
  v_body := v_body || 
    '  </cols>'||xls.newline||
    '  <sheetData>'||xls.newline;     
  call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 1, 1, v_body);
  call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 3, 1, '  </sheetData>'||xls.newline);
  if xls.cells_merge_len>0 then
    call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 3, 9, '  <mergeCells count="1">'||xls.newline);  
    call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 5, 1, '  </mergeCells>'||xls.newline);
  end if;
  call pgxls._add_file_subpart(xls.id, xls.sheet_file_name, 9, 9,
    '  <printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>'||xls.newline||
    '  <pageMargins'||xls.newline||
    '    left   = "'||v_page.margin_left||'"'||xls.newline||
    '    right  = "'||v_page.margin_right||'"'||xls.newline||
    '    header = "'||case when v_page.header_text is not null then v_page.margin_top else 0.000 end||'"'||xls.newline||
    '    top    = "'||case when v_page.header_text is not null then round(v_page.margin_top+(0.035+v_page.header_font_size/65.0)+0.050,3) else v_page.margin_top end||'"'||xls.newline||
    '    bottom = "'||v_page.margin_bottom||'"'||xls.newline||
    '    footer = "'|| 0.000 ||'"'||xls.newline||
    '  />'||xls.newline||
    '  <pageSetup'||xls.newline||
    '    paperSize="'||case when v_page.paper_format='A3' then 8 when v_page.paper_format='A4' then 9 else 11 end||'"'||xls.newline||
    '    orientation="'||v_page.paper_orientation||'"'||xls.newline||     
    '    scale="100" firstPageNumber="1" fitToWidth="1" fitToHeight="1" pageOrder="downThenOver" blackAndWhite="false" draft="false" cellComments="none" useFirstPageNumber="true" horizontalDpi="300" verticalDpi="300" copies="1"'||xls.newline||
    '  />'||xls.newline||    
    '  <headerFooter differentFirst="false" differentOddEven="false">'||xls.newline||    
    '    '||xmlelement(name "oddHeader", 
                       '&'||case when v_page.header_alignment='left' then 'L' when v_page.header_alignment='right' then 'R' else 'C' end||
                       '&"'||v_page.header_font_name||'"'||
                       '&'||v_page.header_font_size||
                       v_page.header_text
                           
            )||xls.newline||
    '    <oddFooter/>'||xls.newline||
    '  </headerFooter>'||xls.newline||
    '</worksheet>'
  );
end
$$;

create or replace procedure pgxls._build_file$xl_shared_strings(inout xls pgxls.xls) language plpgsql as $$
declare
  v_file_name varchar := 'xl/sharedStrings.xml'; 
begin
  call pgxls._add_file_subpart(xls.id, v_file_name, 1, 1,   		
    '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'||xls.newline|| 
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'||xls.strings_len||'" uniqueCount="'||xls.strings_len||'">'||xls.newline
  );
  call pgxls._add_file_subpart(xls.id, v_file_name, 9, 1, 
    '</sst>'
  );  		
 end
$$;
  
create or replace procedure pgxls._build_file$xl_styles(inout xls pgxls.xls) language plpgsql as $$
declare
  v_body text;
  v_styles_len int := array_length(xls.styles,1);
  v_border_side varchar;
  v_border_line pgxls.border_line;
begin	
  v_body :=    
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||xls.newline||
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'||xls.newline||
    '  <numFmts count="'||array_length(xls.formats,1)||'">'||xls.newline;
  for v_format in array_lower(xls.formats,1)..array_upper(xls.formats,1) loop
    v_body := v_body||'    '|| xmlelement(name "numFmt", xmlattributes(v_format as "numFmtId", xls.formats[v_format].code as "formatCode"))||xls.newline;
  end loop;
  v_body := v_body ||
    '  </numFmts>'||xls.newline||
    '  <fonts count="'||array_length(xls.fonts,1)||'">'||xls.newline;
  for v_font in 0..array_length(xls.fonts,1)-1 loop
    v_body := v_body||
      '    <font>'||xls.newline||
      '      <name val="'||xls.fonts[v_font].name||'"/>'||xls.newline||
      '      <sz val="'||xls.fonts[v_font].size||'"/>'||xls.newline||
      '      <b val="'||xls.fonts[v_font].bold||'" />'||xls.newline||
      '      <i val="'||xls.fonts[v_font].italic||'" />'||xls.newline||
      '      <strike val="'||xls.fonts[v_font].strike||'" />'||xls.newline;
    if xls.fonts[v_font].underline then
      v_body := v_body||'<u val="single" />'||xls.newline;
    end if;
    if length(xls.fonts[v_font].color)>=6 then
      v_body := v_body||'      <color rgb="FF'||xls.fonts[v_font].color||'" />'||xls.newline;     
    end if;  
    v_body := v_body||'    </font>'||xls.newline;
  end loop;
  v_body := v_body||    
    '  </fonts>'||xls.newline||
    '  <fills count="'||(2+coalesce(array_length(xls.fills,1),0))||'">'||xls.newline||
    '    <fill><patternFill patternType="none"/></fill>'||xls.newline||
    '    <fill><patternFill patternType="gray125"/></fill>';
  if xls.fills is not null then
    for v_fill in array_lower(xls.fills,1)..array_upper(xls.fills,1) loop
      v_body := v_body||'    <fill><patternFill patternType="solid"><fgColor rgb="FF'||xls.fills[v_fill].foreground_color||'"/></patternFill></fill>';
    end loop;
  end if; 
  v_body := v_body||   
    '  </fills>'||xls.newline||
    '  <borders count="'||array_length(xls.borders,1)||'">'||xls.newline;
  for v_border in 0..array_length(xls.borders,1)-1 loop
    v_body := v_body||
      '    <border diagonalUp="false" diagonalDown="false">'||xls.newline||
      '      <left '||case when  xls.borders[v_border].left_ != 'none' then 'style="'||xls.borders[v_border].left_||'" ' else '' end||'/>'||xls.newline||
      '      <right '||case when  xls.borders[v_border].right_ != 'none' then 'style="'||xls.borders[v_border].right_||'" ' else '' end||'/>'||xls.newline||
      '      <top '||case when  xls.borders[v_border].top != 'none' then 'style="'||xls.borders[v_border].top||'" ' else '' end||'/>'||xls.newline||
      '      <bottom '||case when  xls.borders[v_border].bottom != 'none' then 'style="'||xls.borders[v_border].bottom||'" ' else '' end||'/>'||xls.newline||
      '      <diagonal/>'||xls.newline||
      '    </border>'||xls.newline;
  end loop;
  v_body := v_body|| 
    '  </borders>'||xls.newline||
    '  <cellStyleXfs count="1">'||xls.newline||
    '    <xf numFmtId="100" fontId="0" fillId="0" borderId="0" applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="true">'||xls.newline||
    '      <alignment horizontal="general" vertical="bottom" textRotation="0" wrapText="false" indent="0" shrinkToFit="false"/>'||xls.newline||
    '      <protection locked="true" hidden="false"/>'||xls.newline||
    '    </xf>'||xls.newline||
    '  </cellStyleXfs>'||xls.newline||
    '  <cellXfs count="'||v_styles_len||'">'||xls.newline;
  for v_style in 0..v_styles_len-1 loop
    v_body := v_body ||
      '    <xf numFmtId="'||xls.styles[v_style].format||'" fontId="'||xls.styles[v_style].font||'" fillId="'||xls.styles[v_style].fill||'" borderId="'||xls.styles[v_style].border||'" xfId="0" applyFont="true" applyBorder="true" applyAlignment="true" applyProtection="false">'||xls.newline||
      '      <alignment horizontal="'||xls.styles[v_style].alignment_horizontal||'" indent="'||xls.styles[v_style].alignment_indent||'" vertical="'||xls.styles[v_style].alignment_vertical||'" textRotation="0" wrapText="'||xls.styles[v_style].alignment_text_wrap||'" shrinkToFit="false"/>'||xls.newline||
      '      <protection locked="true" hidden="false"/>'||xls.newline||
      '    </xf>'||xls.newline;
  end loop;
  v_body := v_body ||
    '  </cellXfs>'||xls.newline||
    '  <cellStyles count="1">'||xls.newline||
    '    <cellStyle name="Normal" xfId="0" builtinId="0"/>'||xls.newline||
    '  </cellStyles>'||xls.newline||
    '</styleSheet>';
  call pgxls._add_file_subpart(xls.id, 'xl/styles.xml', 1, 1, v_body);
end
$$;

create or replace procedure pgxls._build_file$xl_workbook(inout xls pgxls.xls) language plpgsql as $$
declare
  v_file_name varchar := 'xl/workbook.xml';
begin
  call pgxls._add_file_subpart(xls.id, v_file_name, 1, 1,
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'||xls.newline||    
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'||xls.newline||
    '  <fileVersion appName="PGXLS version '||pgxls.pgxls_version()||'" />'||xls.newline||
    '  <workbookPr backupFile="false" showObjects="all" date1904="false"/>'||xls.newline||
    '  <workbookProtection/>'||xls.newline||
    '  <bookViews>'||xls.newline||
    '    <workbookView showHorizontalScroll="true" showVerticalScroll="true" showSheetTabs="true" xWindow="0" yWindow="0" windowWidth="16384" windowHeight="8192" tabRatio="500" firstSheet="0" activeTab="0"/>'||xls.newline||
    '  </bookViews>'||xls.newline||
    '  <sheets>'||xls.newline
  );
  call pgxls._add_file_subpart(xls.id, v_file_name, 3, 1,
    '  </sheets>'||xls.newline||
    '  <definedNames>'||xls.newline
  );
  call pgxls._add_file_subpart(xls.id, v_file_name, 9, 1,
    '  </definedNames>'||xls.newline||    
    '  <calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/>'||xls.newline||
    '</workbook>'
  );     
end
$$;

create or replace procedure pgxls._build_file$content_types(inout xls pgxls.xls) language plpgsql as $$
declare
  v_body text; 
begin
  v_body :=
    '<?xml version="1.0" encoding="UTF-8"?>'||xls.newline||	
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'||xls.newline||
    '  <Default Extension="xml" ContentType="application/xml"/>'||xls.newline||
    '  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'||xls.newline||
    '  <Default Extension="png" ContentType="image/png"/>'||xls.newline||
    '  <Default Extension="jpeg" ContentType="image/jpeg"/>'||xls.newline||
    '  <Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'||xls.newline||
    '  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'||xls.newline||
    '  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'||xls.newline;
  for v_sheet in 1..xls.sheets_len loop   
    v_body := v_body ||
      '  <Override PartName="/xl/worksheets/sheet'||v_sheet||'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'||xls.newline;    
  end loop;     
  v_body := v_body ||
    '  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />'||xls.newline|| 
    '  <Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'||xls.newline||
    '  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'||xls.newline||
    '  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'||xls.newline||
    '</Types>';
  call pgxls._add_file_subpart(xls.id, '[Content_Types].xml', 1, 1, v_body);   
end
$$;

create or replace procedure pgxls._zip_add_int(inout zip bytea, bytes int, value bigint) language plpgsql as $$
declare 
  value_text varchar := '\x';
begin
  for i in 1..bytes loop
    value_text := value_text || lpad(to_hex(value & 255),2,'0');
    value := value >> 8;
  end loop;
  zip := zip || (value_text)::bytea; 
end
$$;

create or replace function pgxls._zip_utf8_bytea(value text) returns bytea language plpgsql as $$
declare
  server_encoding varchar := (select pg_encoding_to_char(encoding) from pg_database where datname=current_database());
  value_binary bytea := decode(replace(value, '\', '\\'), 'escape');
begin
  if server_encoding='UTF8' then return value_binary; end if;
  return convert(value_binary, '"'||server_encoding||'"', 'UTF8');
end;
$$;

create or replace procedure pgxls._zip_build(inout xls pgxls.xls) language plpgsql as $$
declare
  v_file record;
  v_file_count int := 0;
  v_zip_len bigint := 0;
  v_cd_len int := 0;
  v_zip_cde bytea; 
  v_datetime bigint;
begin
  call pgxls._trace(xls, '_zip_build', 'started'); 
  v_datetime := ((extract(year from xls.datetime)-1980)::int<<9) | (extract(month from xls.datetime)::int<<5) | extract(day from xls.datetime)::int; 
  v_datetime := (v_datetime<<16) | ((extract(hour from xls.datetime))::int<<11) | ((extract(minute from xls.datetime))::int<<5)  | ((extract(seconds from xls.datetime)/2)::int);	
  for v_file in (select name from (select distinct name from pgxls_temp_file where xls_id=xls.id) f order by name collate "C") loop
    call pgxls._zip_build_file(xls.id, v_file.name, v_datetime, v_file_count, v_zip_len, v_cd_len);
  end loop;
  v_zip_cde:='\x504B0506'::bytea;  -- signature
  call pgxls._zip_add_int(v_zip_cde, 2, 0); -- diskNumber
  call pgxls._zip_add_int(v_zip_cde, 2, 0); -- startDiskNumber
  call pgxls._zip_add_int(v_zip_cde, 2, v_file_count); -- numberCentralDirectoryRecord  
  call pgxls._zip_add_int(v_zip_cde, 2, v_file_count); -- totalCentralDirectoryRecord
  call pgxls._zip_add_int(v_zip_cde, 4, v_cd_len); -- sizeOfCentralDirectory  
  call pgxls._zip_add_int(v_zip_cde, 4, v_zip_len); -- centralDirectoryOffset
  call pgxls._zip_add_int(v_zip_cde, 2, 0); -- commentLength
  insert into pgxls_temp_file(xls_id, name, part, subpart, body) values (xls.id, '~~zip_cde', 1, 1, v_zip_cde); 
end
$$;

create or replace procedure pgxls._zip_build_file(xls_id int, file_name varchar, datetime bigint, inout file_count int, inout zip_len bigint, inout cd_len int) language plpgsql as $$
declare
  v_xls_id int := xls_id;
  v_file_subpart record;
  v_zip bytea;
  v_offset bigint := zip_len;
  v_len bigint := 0;
  v_crc bigint := 4294967295;
begin
  file_count := file_count+1;
  for v_file_subpart in (select body from pgxls_temp_file f where f.xls_id=v_xls_id and f.name=file_name order by part, subpart) loop
    v_len := v_len+length(v_file_subpart.body);
    for i in 0..length(v_file_subpart.body)-1 loop
        v_crc = (v_crc # get_byte(v_file_subpart.body, i))::bigint;
        for j in 1..8 loop
            v_crc := ((v_crc >> 1) # (3988292384 * (v_crc & 1)))::bigint;
        end loop;
    end loop;       
  end loop;
  v_crc := v_crc # 4294967295;
  v_zip:='\x504B0304'::bytea;
  call pgxls._zip_add_int(v_zip, 2, 10);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 2, 0);    
  call pgxls._zip_add_int(v_zip, 4, datetime); 
  call pgxls._zip_add_int(v_zip, 4, v_crc);
  call pgxls._zip_add_int(v_zip, 4, v_len);
  call pgxls._zip_add_int(v_zip, 4, v_len);
  call pgxls._zip_add_int(v_zip, 2, length(file_name));
  call pgxls._zip_add_int(v_zip, 2, 0);
  v_zip := v_zip || pgxls._zip_utf8_bytea(file_name); 
  insert into pgxls_temp_file(xls_id, name, part, subpart, body) values (v_xls_id, file_name, 0, 1, v_zip);
  zip_len:=zip_len+v_len+length(v_zip); 
  v_zip:='\x504B0102'::bytea;
  call pgxls._zip_add_int(v_zip, 2, 10);  
  call pgxls._zip_add_int(v_zip, 2, 10);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 4, datetime);
  call pgxls._zip_add_int(v_zip, 4, v_crc);
  call pgxls._zip_add_int(v_zip, 4, v_len);
  call pgxls._zip_add_int(v_zip, 4, v_len);
  call pgxls._zip_add_int(v_zip, 2, length(file_name));
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 2, 0);
  call pgxls._zip_add_int(v_zip, 4, 0);
  call pgxls._zip_add_int(v_zip, 4, v_offset);
  v_zip := v_zip || pgxls._zip_utf8_bytea(file_name);
  insert into pgxls_temp_file(xls_id, name, part, subpart, body) values (v_xls_id, '~zip_cdf', file_count, 1, v_zip); 
  cd_len := cd_len+length(v_zip);
end;	
$$;

grant usage on schema pgxls to public;
grant execute on all functions in schema pgxls to public;
grant execute on all procedures in schema pgxls to public;


