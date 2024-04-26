Sub MergeExcelFiles()
    Dim fnameList, fnameCurFile As Variant
    Dim countFiles, countSheets As Integer
    Dim wksCurSheet As Worksheet
    Dim wbkCurBook, wbkSrcBook As Workbook
 
	'Here you can specify the file format .xls, .xlsm and so on
    fnameList = Application.GetOpenFilename(FileFilter:="Microsoft Excel CSV (*.xls;*.xlsx;*.xlsm;*.csv),*.xls;*.xlsx;*.xlsm;*.csv", Title:="Choose Excel files to merge", MultiSelect:=True)
 
    If (vbBoolean <> VarType(fnameList)) Then
 
        If (UBound(fnameList) > 0) Then
            countFiles = 0
            countSheets = 0
 
            Application.ScreenUpdating = False
            Application.Calculation = xlCalculationManual
 
            Set wbkCurBook = ActiveWorkbook
 
            For Each fnameCurFile In fnameList
                countFiles = countFiles + 1
 
                Set wbkSrcBook = Workbooks.Open(Filename:=fnameCurFile)
 
                For Each wksCurSheet In wbkSrcBook.Sheets
                    countSheets = countSheets + 1
                    wksCurSheet.Copy after:=wbkCurBook.Sheets(wbkCurBook.Sheets.Count)
                Next
 
                wbkSrcBook.Close SaveChanges:=False
 
            Next
 
            Application.ScreenUpdating = True
            Application.Calculation = xlCalculationAutomatic
 
            MsgBox "Processed " & countFiles & " files" & vbCrLf & "Merged " & countSheets & " worksheets", Title:="Merge Excel files"
        End If
 
    Else
        MsgBox "No files selected", Title:="Merge Excel files"
    End If
End Sub

-----------------------------------------------------------------

Sub SheetNamer()


    Dim i As Integer
    Dim x As Integer
    x = Sheets.Count
    
    For i = 1 To x
    	'inserts in the sheet name into A2
        Sheets(i).Range("A2").Value = Sheets(i).Name
        
    Next i

End Sub


-------------------------------------------------------------------

Sub Combine()
'UpdateByKutools20151029
    Dim i As Integer
    Dim xTCount As Variant
    Dim xWs As Worksheet
    On Error Resume Next
LInput:
    xTCount = Application.InputBox("The number of title rows", "", "1")
    If TypeName(xTCount) = "Boolean" Then Exit Sub
    If Not IsNumeric(xTCount) Then
        MsgBox "Only can enter number", , "Kutools for Excel"
        GoTo LInput
    End If
    Set xWs = ActiveWorkbook.Worksheets.Add(Sheets(1))
    xWs.Name = "Combined"
    Worksheets(2).Range("A1").EntireRow.Copy Destination:=xWs.Range("A1")
    For i = 2 To Worksheets.Count
        Worksheets(i).Range("A1").CurrentRegion.Offset(CInt(xTCount), 0).Copy _
               Destination:=xWs.Cells(xWs.UsedRange.Cells(xWs.UsedRange.Count).Row + 1, 1)
    Next
End Sub
---------------------------------------------------------------------
Sub SheetNamer()

    Dim CareFirst(1 To 11) As String
    Dim i As Integer
    Dim x As Integer
    x = Sheets.Count
    CareFirst(1) = "VCC5.20200511"
    CareFirst(2) = "VCF2.20200511"
    CareFirst(3) = "VCF4.20200511"
    CareFirst(4) = "VCF5.20200511"
    CareFirst(5) = "VCR5.20200511"
    CareFirst(6) = "VHA2.20200511"
    CareFirst(7) = "VHA4.20200511"
    CareFirst(8) = "VHA5.20200511"
    CareFirst(9) = "VPC2.20200511"
    CareFirst(10) = "VRC4.20200511"
    CareFirst(11) = "VRC5.20200511"
    

    For i = 1 To x
        'inserts in the sheet name into A2
        Sheets(i).Name = CareFirst(i)
        Sheets(i).Range("A2").Value = Sheets(i).Name
        
    Next i

End Sub
