Attribute VB_Name = "vbaUtilities"
Dim apiSh As Worksheet, firSh As Worksheet

Public Const apiName = "APIKeySheet"

Public Const timeOutMilSec = 20000

Option Explicit

Sub Code_Run(blRun As Boolean)
    If blRun Then
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
        Application.ScreenUpdating = False
        Application.UserControl = False
    Else
        Application.EnableEvents = True
        Application.Calculation = xlCalculationAutomatic
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.UserControl = True
    End If
End Sub

'Verify that the API sheet exists
Sub Verify_API_Sheet()
    Dim sh As Worksheet
    Dim blShtEx As Boolean
    
    'Check to make sure sheet doesn't already exist
    blShtEx = False
    
    For Each sh In ActiveWorkbook.Sheets
        If sh.name = apiName Then
            blShtEx = True
                
            If sh.Visible = xlSheetVisible Then
                sh.Visible = xlSheetVeryHidden
            End If
        End If
    Next
    
    'If not create and hide
    If Not blShtEx Then
        ActiveWorkbook.Sheets.Add After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        
        Set sh = ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count)
        
        sh.name = apiName
        sh.Visible = xlSheetVeryHidden
    End If

End Sub

Sub Unhide_sheets()
    Dim sh As Worksheet
    Dim WS_count As Long, i As Long
    
    Code_Run True
    
    WS_count = ActiveWorkbook.Worksheets.Count
    ActiveWorkbook.Sheets("APIKeySheet").Cells(1, 1).value = vbNullString
    
    For i = 1 To WS_count
        ActiveWorkbook.Sheets(i).Visible = xlSheetVisible
    Next
    
    For Each sh In ActiveWorkbook.Sheets
        sh.Visible = xlSheetVisible
    Next
    
    Code_Run False
End Sub

Sub A()
    Application.EnableEvents = True
    ActiveWorkbook.Sheets("APIKeySheet").Visible = xlSheetVeryHidden
End Sub

'Code to handle Errors
Sub ErrorHandling(ByVal ProcedureName As String, ByVal AdditionalData As String)

    If AdditionalData <> "" Then
        MsgBox "The application experienced an error with the " & ProcedureName & " (" & AdditionalData & ") procedure."
    Else
        MsgBox "The application experienced an error with the " & ProcedureName & " procedure."
    End If
    
    Err.Clear
    
    'Code_Run False
    
    'Application.OnTime Now() + TimeSerial(0, 0, 0.5), "A"

End Sub

Function Is_MAC() As Boolean
    Dim exVers As Double
    Dim runPC As Boolean

    exVers = val(Application.Version)
    
    If exVers - WorksheetFunction.Floor(exVers, 1) = 0 Then
        runPC = False
    Else
        runPC = True
    End If
    
    Is_MAC = runPC
End Function

Sub Toggle_Change(toggle As Boolean)
    Dim countr As Long, countc As Long
    
    Set firSh = ActiveWorkbook.Sheets(1)
    
    countr = firSh.Rows.Count
    countc = firSh.Columns.Count
    
    'firSh.Cells(countr, countc).Value2 = toggle
    
    'countr = countr
    
End Sub

Function Check_Toggle() As Boolean
    Dim countr As Long, countc As Long
    
    Set firSh = ActiveWorkbook.Sheets(1)
    
    countr = firSh.Rows.Count
    countc = firSh.Columns.Count
    
    If firSh.Cells(countr, countc) Then
        Check_Toggle = True
    Else
        Check_Toggle = False
    End If
    
End Function

'Function To Confirm Date Inputs
Function Check_Date_Input(dateStr As Variant) As Boolean

    If Not IsDate(CStr(dateStr)) Then
        Check_Date_Input = False
        Exit Function
    End If
    
    If CDate(dateStr) < DateSerial(1920, 1, 1) Then
        Check_Date_Input = False
        Exit Function
    End If
    
    Check_Date_Input = True
End Function

'Function To Confirm Time Inputs
Function Check_Time_Input(timeStr As Variant) As Boolean

    If Not IsDate(CStr(timeStr)) Then
        Check_Time_Input = False
        Exit Function
    End If
    
    If CDate(timeStr) > 1 Then
        Check_Time_Input = False
        Exit Function
    End If
    
    Check_Time_Input = True
End Function

Sub Error_Message(error_msg As String)
    If error_msg <> vbNullString Then
        MsgBox error_msg, vbCritical, "RavenPack"
    End If
End Sub

Function Response_Error_Handle(resp As WebResponse, Optional rngCall As Range) As String
    Dim var As Variant
    Dim outMsg As String
            
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    apiSh.Unprotect

    'Debug.Print resp.Data("errors")(1)("reason")
    

    If Not resp.Data Is Nothing Then
        'String returned indicating one error
        If VarType(resp.Data("errors")) = 8 Then
            If Not rngCall Is Nothing Then
                outMsg = "Query in cell: " & rngCall.Address & " had the following errors: " & vbCrLf & CStr(resp.Data("errors"))
            Else
                outMsg = "Query had the following errors: " & vbCrLf & CStr(resp.Data("errors"))
                Debug.Print resp.Data("errors")
            End If
        'else if dictionary of errors returned
        Else
            For Each var In resp.Data("errors")
                If outMsg = vbNullString Then
                    If Not rngCall Is Nothing Then
                        outMsg = "Query in cell: " & rngCall.Address & " had the following errors: " & vbCrLf & _
                        "Error: " & var("type") & " " & var("reason") & " " & var("paramater_name")
                    Else
                        outMsg = "Query had the following errors: " & vbCrLf & _
                        "Error: " & var("type") & " " & var("reason") & " " & var("paramater_name")
                    End If
                Else
                    outMsg = outMsg & vbCrLf & "Errors: " & var("type") & " " & var("reason") & " " & var("paramter_name")
                End If
            Next
        End If
    Else

    End If

    If outMsg = vbNullString And resp.StatusDescription <> vbNullString Then
        outMsg = resp.StatusDescription
    End If
    '
    Response_Error_Handle = outMsg
    
End Function


