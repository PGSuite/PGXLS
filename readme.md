## PGXLS | Creating Excel tables in PostgreSQL

### [Download](https://pgxls.org/en/download/) ###
### [Documentation](https://pgxls.org/en/documentation/) ### 
### [Example](https://pgxls.org/en/#examples-full) ### 

### Description ###

Tool PGXLS is schema with stored procedures for create file (bytea type) in Excel format (.xlsx).
Implemented dependence format on data type, conversion SQL query into sheet with autoformat and more.


### Basic procedures ###
  
*   **pgxls.create** - create document
  
*   **pgxls.add_row** - add row
  
*   **pgxls.set_cell_value** - set cell value
  
*   **pgxls.get_file** - build and get file


### Important qualities ### 

*   **Large files** - data row by row inserted into temporary table, which not requires memory. Separate function is implemented to get large file
*   **Auto-format** - for each column, format is configured depending on data type
*   **SQL queries** - it is possible to add sheet with the results of SQL query
*   **Styles** - for columns and cells, support format, font, border, fill and alignment
*   **Parallelism** - possible to create several files in parallel in one session
