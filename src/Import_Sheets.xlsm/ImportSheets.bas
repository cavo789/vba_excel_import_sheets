Attribute VB_Name = "ImportSheets"
Public Sub DoIt()

Dim sFileName As String
Dim wb As Workbook, wbNew As Workbook
Dim shSource As Worksheet, shTarget As Worksheet

    sFileName = shParams.Range("rngFileName").Value
    
    If (Dir(sFileName) = "") Then
       MsgBox "Sorry this file doesn't exists", vbCritical
       Exit Sub
    End If
    
    Set wbNew = Application.Workbooks.Add
    
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = True
    
    Set wb = Workbooks.Open(Filename:=sFileName, UpdateLinks:=False, ReadOnly:=True)
    
    If (wb Is Nothing) Then
       MsgBox "Sorry the file seems to be corrupt, Workbooks.Open() has failed", vbCritical
       Exit Sub
    End If
    
    For Each sh In wb.Worksheets
    
        Application.StatusBar = "Import sheet [" & sh.Name & "] from [" & wb.Name & "]"
          
        On Error Resume Next
        
        Call sh.Copy(Before:=wbNew.Worksheets(wbNew.Worksheets.Count))
        
        If Err.Number <> 0 Then
            MsgBox "Error " & sh.Name & " : " & Err.Description, vbCritical
            Err.Clear
        End If
        
    Next
    
    Call wb.Close(SaveChanges:=False)
    Set wb = Nothing
    
    wbNew.Activate
    
    MsgBox "Sheets copied from " & sFileName & " to the current workbook", vbInformation
    
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    
End Sub
