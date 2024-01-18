# Simple-Data-Miner
Extract data from a database with no SQL knowledge

Simple Data Miner uses Python, Flask, SQLAlchemy and pandas to create a web site where users,
with little or no SQL knowledge, can extract the data they are looking for from a database.
It does this by simplifying the process. Users are presented with a list of "tables",
but they can only pick **one "table"**. Those "tables" are usually views on the data.
Views give database administrators many advantages.
* Database columns can be renamed, with longer, more descriptive, but still SQL compliant names
* Database table joins can be hidden from the users
* Deleted or cancelled records can be filtered out of the returned data
* Derived columns can be added, such as the ratio of two column values

Each "table" has a configuration which includes a more user friendly name for the "table"
plus the maximum number of rows that can be included in the mined extraction, ignoring any aggregation.
That is, the maximum number of rows of data fetched from the database.

The colums can be given even longer, more user friendly, non-SQL compliant names.
Columns that are indexed are flagged, as the user must select and constrain at least one indexed column.
Columns which contain codes, associated with a code/description lookup table can be identified,
with the name of that lookup table and the names of the lookup table's code and description columns.

Once the user has selected their "table" they are presented with a checkbox list
of the columns and asked to select the column they want mined into their extract.
Indexed columns will have an ** (indexed)** appended to their long, user friendly name.
If the user doesn't select at least one indexed column, they will be advised of their
error and asked to try again.

Next the user will be asked if they wish to constrain any of the columns.
If the user doesn't choose to constrain at least one indexed column, they will be advised of their
error and asked to try again.
For each chosen contraint the user will be asked to configure that constraint.

If the column is a code, associated with a lookup table, then the codes and descriptions
from the lookup table will be displayed as a checkbox list and the user asked to select
the codes that they want included in their mined extract.

For other columns, the possible constraints will be displayed, as appropriate to the data type
of the chosen colum. For numbers and dates that is things like "less than a value", "in a range" etc.
For character columns it will be options such as "starts with", "equal to", "contains" etc.
For the chosen constrain type, the user will be asked to enter the parameters for that constraint,
such as a minimum value and a maximum value. The entered data will be checked to ensure that
it matches the data type of the column being constrained.
If there is a datatype error the user will be advised of their error and asked to try again.

Next, the users will be asked if they want to "count" and/or "sum" any columns.
Only columns with data types appropriate for "count()" and "sum()" will be displayed.

Finally, the user will be shown the configures SQL query and asked if this is the query
they wish to run. Thus, users will slowly, but suerly, learn SQL by example.

Finally, the data is extracted into an Excel workbook and download.
The workbook will contain two sheets
* **SQL query** being the SQL query run against the database
* **mined extract** being the extracted data

## Features
* Simple user interface
* Extract is automatically download as an Excel workbook
* Almost all constraints on data are supported
  + For strings (=, !=, startsWith, endsWith, contains, does not contain)
  + For numbers and dates (=, !=, <, <=, >, >=, between two values)
* count(), sum(), avg(), min() and max() aggreagtions are supported
* Mined data can be previewed before being downloaded

## Limitations
The **Simple Data Miner** is "simple" and has such it has limitations. However, in workarounds for most of these limitations.
* Only **AND** is supported. Users can constrain as many columns as they like,
with more than one constraint on each column. However, the resulting SQL is
the AND of all of these constraints. The **OR** construct is not supported.
There are workarounds for most situations where **OR** would be convenient.
  + colA = "x" **OR** colA = "y" - do two extracts. Both extracts can be downloaded as Excel workbooks which can be concatenated.
    - colA = "x"
    - colB = "y" AND colA != "x'
  + colA = "x" **OR** colA = "y" - do two extracts. Both extracts can be downloaded as Excel workbooks which can be concatenated.
    - colA = "x"
    - colA = "y"
  + colA <= "x" **OR** colA >= "y" - [not in a range] - do two extracts. Both extracts can be downloaded as Excel workbooks which can be concatenated.
    - colA <= "x"
    - colA >= "y"
* **"in"** a list is not supported (other than for columns associated with a lookup table).
  + colA in [1, 3, 5, 7 , 9] - do an extract for each of these values.
* Relationships between columns are not supported
  + colA <= colB - there is no workaround. Users can extract all the data into the Excel download and use Excel formulas to create a derived column of "=colA <= colB" and then use a Pivot table to select the required data, but that is  hardly a "workaround".
* Derived columns are not supported. Users can only extract columns that are in the "table". Users cannot create a new column using formulas combining data from other columns. Users can extract all the data into the Excel download and use Excel formulas to create the derived column, but that is  hardly a "workaround".