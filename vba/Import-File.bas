Attribute VB_Name = "Module1"
Option Explicit

Sub ImportData()

Dim sourcepath As String
Dim archivepath As String
Dim filename As String

sourcepath = "D:\VBA Macro Excel\Source Data\"
archivepath = "D:\VBA Macro Excel\Archive\"

filename = Dir(sourcepath & "*.csv")

Do While filename <> ""
    Workbooks.Open (sourcepath & filename)
    Range("A1").CurrentRegion.Offset(1, 0).Copy _
    ThisWorkbook.Sheets("RawData").Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
    Workbooks(filename).Close False
    
    Name sourcepath & filename As archivepath & filename
    
    filename = Dir
Loop

End Sub
