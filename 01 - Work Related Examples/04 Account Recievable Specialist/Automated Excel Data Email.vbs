Sub emailNurse()
 
 ' Before you run this module, copy the following function and paste it in Cell "U2"
 ' =CONCATENATE(TRIM(LEFT(MID(P2,FIND(",",P2)+2,LEN(P2)-FIND("(#",P2)),1)),TRIM(LEFT(P2,FIND(",",P2)-1)),"@google.ca")
 ' Afterwards, Paste any other emails you want to cc in "V" coloumn
 '
 '
 '
 ' Now run the module using the play button in the ribbon above
 
 Dim MainEmail As String
 Dim SupportEmail As String
 Dim SupervisorEmail As String
 Dim SubjectLine As String
 Dim rowww As Integer
 Dim Mail_Object, Email_Subject
 Dim o As Variant
 Dim lr As Long
 lr = Cells(Rows.Count, "A").End(xlUp).Row 'last row counter
 Set Mail_Object = CreateObject("Outlook.Application")
 For rowww = 2 To 3 'lr
    If Range("R" & rowww).Value = "Inform*" Then
        With Mail_Object.CreateItem(o)
                .Subject = "Rejected Visit: " & Range("E" & rowww).Value
                .To = Range("U" & rowww).Value
                .cc = Range("V" & rowww).Value
                .HTMLbody = "<font size=""10pt"" face=""Calibri"" color="""">" & "Hello," & "<br><br>" & "Please f/up with the rejections for:" & "<br><br>" & "Client: " & Range("E" & rowww).Value & "<br><br>" & "Rejected Date: " & Range("F" & rowww).Value & "<br>" & "<font size=""10pt"" face=""Calibri"" color=""red"">" & "<B>LHIN authorized:</B> " & "<br>" & "<font size=""10pt"" face=""Calibri"" color=""black"">" & "<B>" & Range("W" & rowww).Value & "</b>" & "<br><br><br><br><br><br>" & "Thank you,"
                '.Send
                .display 'disable display and enable send to send automatically
        End With
        Range("R" & rowww).Value = "Email sent to Nurse/CC"
    End If
 Next rowww

End Sub

