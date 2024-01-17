# Simple-Data-Miner
Extract data from a database with no SQL knowledge

Simple Data Miner uses Python, Flask, SQLAlchemy and pandas to create a web site where users,
with little or no SQL knowledge, can extract the data they are looking for.
It does this by simplifying the process. Users are presented with a list of "tables",
but they can only pick **one "table"**. Those "tables" are usually views on the data.
Views give database administrators many advantages.
* Database columns can be renamed, with longer, more descriptiong, but still SQL compliant names
* Database table joins can be hidden from the users
* Deleted or cancelled records can be filtered out of the returned data
* Derived columns can be added, such as the ratio of two column values

Each "table" has a configuration which includes a more user friendly name for the "table"
plus the maximum number of rows that can be returned.
The colums can be given even longer, more user friendly, non-SQL compliant names.
Columns that are indexed are flagged, as the user must select at least one indexed column.
Columns which contain codes, associated with a code/description lookup table can be identified,
with the name of that lookup table and the names of it's code and description columns.

Once the user has selected their "table" they are presented with a checkbox list
of the columns and asked to select the column they want mined into their extract.
Indexed columns will have an **\*** appended to their long, user friendly name.
If the user doesn't select at least one indexed column, they will be advised of their
error and asked to try again.

Next the user will be asked if they wish to constrain any of the columns.
For each chosen contraint the user will be asked to configure that constraint.

If the column is a code, associated with a lookup table, then the codes and descriptions
from the lookup table will be displayed as a checkbox list and the user asked to select
the codes that they want included in their mined extract.

For other columns, the possible constraints will be displayed, as appropriate to the data type
of the chosen colum. Ffor numbers and dates that is things like "less than a value", "in a range" etc.
For character columns it will be options such as "starts with", "equal to", "contains" etc.
For the chosen constrain type, the user will be asked to enter the parameters for that constraint,
such as a minimum value and a maximum value. The entered data will be checke to ensure that
it matches the data type of the column being constrained.

Next, the users will be asked if they want to "count" or "sum" any columns.
Only columns with data types appropriate for "count()" and "sum()" will be displayed
and the user will be offered the choice of "count", "sum" or "count and sum".

Finally, the user will be shown the configures SQL query and asked if this is the query
they wish to run. Thus, users will slowly, but suerly, learn SQL by example.

Finally, the data is extracted into an Excel workbook which the user can download.