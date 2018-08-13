Sub DataMerge_NIR()
Dim uploadfile As Variant
Dim uploader As Workbook
Dim CurrentBook As Workbook
 
Set CurrentBook = ActiveWorkbook
uploadfile = Application.GetOpenFilename()
    If uploadfile = "False" Then
        Exit Sub
    End If
   
Workbooks.Open uploadfile
Set uploader = ActiveWorkbook
With uploader
    Application.CutCopyMode = False
    Range("E3:E200").Copy
End With
uploader.Activate
Sheets(1).Name = "Sheet1"
Sheets("Sheet1").Range("Z2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
End Sub
