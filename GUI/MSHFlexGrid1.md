referensi:
http://www.ucalgary.ca/cpsc
http://cariprogram.blogspot.com/
http://www.maranatha.edu/

```
' menghapus semua data pada MSFlexGrid
MSHFlexGrid1.Clear
```

```
' menambah satu baris data
MSHFlexGrid1.AddItem ""
```

```
' menghapus satu baris data
' minimal 1 baris tidak dapat dihapus
MSHFlexGrid1.RemoveItem 1
```

```
' menghapus semua baris
MSHFlexGrid1.Rows = 0
```
