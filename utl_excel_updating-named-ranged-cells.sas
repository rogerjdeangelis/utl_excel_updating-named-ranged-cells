Updating named ranged cells in Excel using SAS 9.4

Problem: Change Alfreds age to 99 in named range class

Multiple solutions (you need classic SAS for this?)


INPUT
=====
    d:/xls/class.xlsx  with sheet name class and named range class


    WORKBOOK d:/xls/class.xlsx with sheet class (you can use [sheet1])

   d:/xls/class.xlsx
      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    14      |    69      |  112.5     |
      +------------+------------+------------+------------+------------+
   3  | BARBARA    |    F       |    13      |    58      |  101.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

   [CLASS]


PROCESSES
=======

 SAS  (age can be macro variable or in anothe SAS or excel named range)

   libname xel "d:/xls/class.xlsx" scan_text=no;
   data xel.class;
      modify xel.class;
      age=99;
      where name= 'Alfred';
   run;quit;
   libname xel clear;

 SAS/Passthru

    * this does it in place;
    proc sql dquote=ansi;
       connect to excel as excel(Path="d:/xls/class.xlsx");
       execute(
         update [class]
         set age=88
         where name="Alfred"
       ) by excel;
       disconnect from excel;
    Quit;

  Python (see below)

OUTPUT
=====
    d:/xls/class.xlsx  with sheet name class and named range class


    WORKBOOK d:/xls/class.xlsx with sheet class (you can use [sheet1])

   d:/xls/class.xlsx
      +----------------------------------------------------------------+
      |     A      |    B       |     C      |    D       |    E       |
      +----------------------------------------------------------------+
   1  | NAME       |   SEX      |    AGE     |  HEIGHT    |  WEIGHT    |
      +------------+------------+------------+------------+------------+
   2  | ALFRED     |    M       |    99      |    69      |  112.5     |   ** note age changed to 99;
      +------------+------------+------------+------------+------------+
   3  | BARBARA    |    F       |    13      |    58      |  101.5     |
      +------------+------------+------------+------------+------------+
       ...
      +------------+------------+------------+------------+------------+
   20 | WILLIAM    |    M       |    15      |   66.5     |  112       |
      +------------+------------+------------+------------+------------+

   [CLASS]

  There are many ways to do this
    1. Using macro variables and explicit passthru to excel
    2. Creating a temp table with updates and using passthru to jupdate using the two tables
    3. Bring the nameded range into sas update and send back
    4. Using Python openpyxl ( I like this the best see below )

https://goo.gl/dJaQsG
https://communities.sas.com/t5/Base-SAS-Programming/Updating-named-ranged-cells-in-Excel-using-SAS-9-4/m-p/416722

 *                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;
* create named range class;
libname xel "d:/xls/class.xlsx" scan_text=no;
data xel.class;
  set sashelp.class;
run;quit;
libname xel clear;


*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __  ___
/ __|/ _ \| | | | | __| |/ _ \| '_ \/ __|
\__ \ (_) | | |_| | |_| | (_) | | | \__ \
|___/\___/|_|\__,_|\__|_|\___/|_| |_|___/

;

libname xel "d:/xls/class.xlsx" scan_text=no;
data xel.class;
   modify xel.class;
   age=99;
   where name= 'Alfred';
run;quit;
libname xel clear;

proc sql dquote=ansi;
   connect to excel as excel(Path="d:/xls/class.xlsx");
   execute(
     update [class]
     set age=88
     where name="Alfred"
   ) by excel;
   disconnect from excel;
Quit;


/* T1004450 SAS Forum: Python Update excel "rectangle" within a named range without using column names

Easier with column names?

HAVE
====

Exel sheet

d:/xls/class.xlsx

------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |F           |13            |
|------------+------------+--------------+
|Barbara     |F           |13            |
|------------+------------+--------------+
|Carol       |F           |14            |
|------------+------------+--------------+
|Henry       |M           |14            |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+

[CLASS]


Up to 40 obs SD1.CLASS total obs=3

Obs    NAME       SEX    AGE

 1     Alice       A     112
 2     John        J     111
 3     William     W     114

WANT
====

------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |                           |
|------------+                           +
|Barbara     |    UPDATE THIS INPLACE    |
|------------+                           +
|Carol       |                           |
|------------+------------+--------------+
|Henry       |M           |14            |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+


------------------------------------------
|    A       |     B      |    C         |
|----------------------------------------+
|NAME        |SEX         |AGE           |
|------------+------------+--------------|
|Alfred      |M           |14            |
|------------+------------+--------------+
|Alice       |A           |112           |
|------------+------------+--------------+
|Barbara     |B *updated  |111 *updated  |
|------------+------------+--------------+
|Carol       |C *updated  |114 *updated  |
|------------+------------+--------------+
|Henry       |M *updated  |14  *updated   |
|------------+------------+--------------+
|James       |M           |12            |
|------------+------------+--------------+
|Jane        |F           |12            |
|------------+------------+--------------+
|Janet       |F           |15            |
|------------+------------+--------------+
|Jeffrey     |M           |13            |
|------------+------------+--------------+

[CLASS]

WORKING CODE
============

     PYTHON CODE

       for r_idx in range(3):;
           for c_idx in range(2):;
                c=c_idx+1;
                r=r_idx+1;
                ws.cell(row=r_idx+3, column=c_idx+2,value=clas.iloc[r-1,c-1]);

*                _              _       _
 _ __ ___   __ _| | _____    __| | __ _| |_ __ _
| '_ ` _ \ / _` | |/ / _ \  / _` |/ _` | __/ _` |
| | | | | | (_| |   <  __/ | (_| | (_| | || (_| |
|_| |_| |_|\__,_|_|\_\___|  \__,_|\__,_|\__\__,_|

;

options validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.class;
  retain name sex age;
  set sashelp.class(keep=name age where=(name in ('Alice','Barbara','Carol')));
  sex=substr(name,1,1);
  age=age+99;
  keep name sex age;
run;quit;
libname sd1 clear;


%utlfkil(d:/xls/class.xlsx);
libname xel "d:/xls/class.xlsx";
data xel.class;
  set sashelp.class(keep=name sex age);
run;quit;
libname xel clear;

/*
Up to 40 obs from sd1.class total obs=3

Obs    SEX    AGE

 1      A     112
 2      J     111
 3      W     114
*/

*          _       _   _
 ___  ___ | |_   _| |_(_) ___  _ __
/ __|/ _ \| | | | | __| |/ _ \| '_ \
\__ \ (_) | | |_| | |_| | (_) | | | |
|___/\___/|_|\__,_|\__|_|\___/|_| |_|

;

* this works;
%utl_submit_py64old("
from openpyxl.utils.dataframe import dataframe_to_rows;
from openpyxl import Workbook;
from openpyxl import load_workbook;
from sas7bdat import SAS7BDAT;
with SAS7BDAT('d:/sd1/class.sas7bdat') as m:;
.   clas = m.to_data_frame();
wb = load_workbook(filename='d:/xls/class.xlsx', read_only=False);
ws = wb.get_sheet_by_name('class');
rows = dataframe_to_rows(clas);
for r_idx in range(3):;
.   for c_idx in range(2):;
.        c=c_idx+1;
.        r=r_idx+1;
.        ws.cell(row=r_idx+3, column=c_idx+2,value=clas.iloc[r-1,c-1]);
wb.save('d:/xls/class.xlsx');
");






