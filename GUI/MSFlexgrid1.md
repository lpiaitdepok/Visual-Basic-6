'Pilih MSFlexgrid1 tersebut dengan meng-klik-nya.
'Klik pada DataSource property (pada properties window) dan rubah captionnya menjadi Data1

```
If Not Dir(App.Path & "\NWIND.mdb") = "" Then

Data1.DatabaseName = App.Path & "\NWIND.mdb"

Data1.RecordSource = "SELECT * FROM Customers"

End If
```
