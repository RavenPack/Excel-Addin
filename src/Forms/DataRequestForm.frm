VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DataRequestForm 
   Caption         =   "RavenPack Analytics - Data Request"
   ClientHeight    =   4670
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6825
   OleObjectBlob   =   "DataRequestForm.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "DataRequestForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim apiSh As Worksheet

Private Sub Userform_Initialize()
    Dim api_key As String
    
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    save_api_key
    api_key = apiSh.Cells(1, 1)
    
    Me.api_key_box.value = api_key
    
End Sub

Private Sub Pasteclip_button_Click()
    Dim DataObj As MSForms.DataObject
    Set DataObj = New MSForms.DataObject
    
    On Error GoTo ErrorHandle
    
    ' Read from the clipboard, removing whitespace
    DataObj.GetFromClipboard
    
    ' Take the last item from the clipboard
    clip_text = Trim(DataObj.GetText(1))
    
    ' Remove any extra newline characters
    clip_text = Replace(Replace(clip_text, Chr(10), ""), Chr(13), "")
    
    ' Set the value of the dataset form
    DataRequestForm.dataset_uuid_box.value = clip_text
    
    Exit Sub
    
ErrorHandle:
    If Err.Number <> -2147210493 Then
        MsgBox Err.Description
        Err.Clear
    Else
        Err.Clear
    End If
End Sub

Private Sub RunDataRequest_Click()

    'GET VARIABLES FROM DatasetForm
    'Check Start Date
    If Not Check_Date_Input(start_date_box.value) Then
        MsgBox "Please supply a start date in the correct format (YYYY-MM-DD)"
        Exit Sub
    End If
    'Check Start Time
    If Not Check_Time_Input(start_date_time_box.value) Then
        MsgBox "Please supply a start time in the correct format (hh:mm:ss)"
        Exit Sub
    End If
    
    'Check End Date
    If Not Check_Date_Input(end_date_box.value) Then
        MsgBox "Please supply an end date in the correct format (YYYY-MM-DD)"
        Exit Sub
    End If
    'Check End Time
    If Not Check_Time_Input(end_date_time_box.value) Then
        MsgBox "Please supply a end time in the correct format (hh:mm:ss)"
        Exit Sub
    End If
    
    start_date = start_date_box.value
    start_date_time = start_date_time_box.value
    end_date = end_date_box.value
    end_date_time = end_date_time_box.value
    dataset_uuid = dataset_uuid_box
    api_key = api_key_box.value
    ' Not sure what this line is for...
    api_key = Format(api_key, text)
    Call DataRequest(start_date, start_date_time, end_date, end_date_time, dataset_uuid, api_key)
End Sub

Private Sub DataRequestForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        DataRequestForm.Hide
    End If
End Sub

