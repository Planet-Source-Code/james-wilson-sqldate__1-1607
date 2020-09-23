<div align="center">

## SQLDate


</div>

### Description

Dates in SQl queries often cause problems, as the date must be in the ANSI format whereas dates brought back can be in a different local format. This function simply returns the date in the required format and save having to type Format(DateString, "mm/dd/yy") every time.
 
### More Info
 
The date to be processed as type DATE.

Example SQl Query-

SQL = "SELECT * from tblTest"

SQL = SQL & " WHERE StartDate = #" & SQLDate(DateToConvert) & "#

A STRING containing the date formatted to the correct criteria.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[James Wilson](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/james-wilson.md)
**Level**          |Unknown
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/james-wilson-sqldate__1-1607/archive/master.zip)





### Source Code

```
Public Function SQLDate(ConvertDate As Date) As String
  SQLDate = Format(ConvertDate, "mm/dd/yyyy")
End Function
```

