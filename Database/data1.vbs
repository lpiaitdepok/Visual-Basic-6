

'make sure the search path to the db is always in the right spot
Data1.DatabaseName = App.Path & "\Student.mdb"

'set up the recordsource for the datasource and flexgrid control
'in this case, it's just a raw SQL query, simple simple.
Data1.RecordSource = "SELECT * FROM STUDENT ORDER BY StudentName"

'End Of File
Data1.Recordset.EOF

'Begin Of file
Data1.Recordset.BOF

'
Data1.Recordset.MoveFirst

'add a new entry to our table.
Data1.Recordset.AddNew !StudentName = txtStudentName.Text
Data1.Update
Data1.Refresh

'delete an entry from the database
Data1.Recordset.Move ( MSFlexgrid1.Row - 1) ' we minus one because row zero is the header row
Data1.Delete
Data1.Refresh
