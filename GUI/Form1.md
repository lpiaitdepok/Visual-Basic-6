```
' font dalam form
Form1.Font = "Courier New"
Form1.FontSize = 10

' menampilkan form1
Form1.Show

' koordinat x
Form1.CurrentX = 0
Form1.CurrentY = 0

' hanya bisa digunakan setelah event form activate
Form1.Print
' contoh
Form1.Print Tab(6); "BIODATA";
Form1.Print Tab(6); "MAHASISWA "; Format(Time, "hh:mm AM/PM");
Form1.Print Tab(2); "==========================================";
Form1.Print Tab(3); "NIM :"; Text1.Text;



' keluar dari form1
Unload Form1
```
