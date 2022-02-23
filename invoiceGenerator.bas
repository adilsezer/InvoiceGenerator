Attribute VB_Name = "InvoiceGenerator"
Sub InvoiceGenerator()

Dim clientid As String: clientid = Trim(UCase(InputBox("Please enter your client ID", "PDF Invoice Macro")))
If clientid = vbNullString Then Exit Sub

Dim macroWb As Workbook: Set macroWb = ThisWorkbook
Dim macroSht As Worksheet: Set macroSht = macroWb.Sheets("Invoice Macro")
Dim templateSht As Worksheet: Set templateSht = macroWb.Sheets(clientid)
Dim detailSht As Worksheet: Set detailSht = ThisWorkbook.Sheets("Invoice Details")

If detailSht.Range("A2") = "" Then MsgBox ("Please fill invoice details in 'Invoice Details' sheet")

Application.DisplayAlerts = False

directoryPath = Application.ThisWorkbook.Path
strFolder = "Created PDF Invoices on " & Replace(Replace(Now(), "/", "."), ":", "-")

If Len(Dir(directoryPath & "/" & strFolder, vbDirectory)) = 0 Then
  MkDir (directoryPath & "/" & strFolder)
End If

Set newTemplateSht = macroWb.Worksheets.Add(After:=macroWb.Worksheets(macroWb.Worksheets.Count), Type:=xlWorksheet)
newTemplateSht.Range("A:Z").NumberFormat = "@"

'''''''''''''''''''''''''PDF FILE STARTS'''''''''''''''''''''''''''''''''''''''''''''''''
With detailSht
    For Each invoiceno In .Range("D2:D" & .Cells(.Rows.Count, "A").End(xlUp).Row)
        
        templateSht.Range("A1:Z40" & LastRow).Copy newTemplateSht.Range("A1")
        Application.Wait Now + TimeValue("0:00:01")

        junk = Array("<", ">", ":", """", "|", "?", "*")
        For Each A In junk
            StrFile = Replace(invoiceno.Value2, A, "")
        Next A
        
        Dim fndList As Variant: fndList = Array("Uniqueremitter", "Uniqueamt", "Uniquecry", "Uniqueinvno", "Uniquedate", "Uniquematter", "Uniquedesc")
        Dim rplcList As Variant: rplcList = Array(invoiceno.Offset(0, -3).Value2, invoiceno.Offset(0, -2).Value2, invoiceno.Offset(0, -1), invoiceno.Value2, invoiceno.Offset(0, 1).Value2, invoiceno.Offset(0, 2), invoiceno.Offset(0, 3))
        
        'Loop through each item in Array lists
          For x = LBound(fndList) To UBound(fndList)
                newTemplateSht.Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
                    Lookat:=xlPart, SearchOrder:=xlByRows, MatchCase:=True, _
                    SearchFormat:=False, ReplaceFormat:=False
          Next x
            
        If Len(Dir(directoryPath & "/" & strFolder & "/" & StrFile & ".pdf")) = 0 Then
            newTemplateSht.ExportAsFixedFormat _
                Type:=xlTypePDF, _
                Filename:=directoryPath & "/" & strFolder & "/" & StrFile, _
                OpenAfterPublish:=False, _
                Quality:=xlQualityStandard, _
                IncludeDocProperties:=True, _
                IgnorePrintAreas:=False
        Else
            MsgBox ("File " & StrFile & ".pdf already exist, skipping this invoice!")
        End If
            
        newTemplateSht.Cells.Delete
        newTemplateSht.Pictures.Delete
    
    Next invoiceno
    newTemplateSht.Delete
End With

''''''''''''''''''''''''''PDF FILE END'''''''''''''''''''''''''''''''''''''''''''''''''
macroSht.Activate

Application.DisplayAlerts = True

MsgBox ("Done! PDF invoices were created in " & strFolder & " folder")

End Sub

