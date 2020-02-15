Dim db as Database
Dim rs as Recordset

Set db = CurrentDb
Set rs = db.OpenRecordset("SELECT [Name] FROM [MSysObjects] WHERE ([Type] =  6);", dbOpenSnapshot, dbForwardOnly)
Do While (Not rs.EOF)
    db.TableDefs.Delete rs.Fields("Name").Value
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
db.Close
Set db =  Nothing

