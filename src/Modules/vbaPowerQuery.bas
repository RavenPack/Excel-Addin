Attribute VB_Name = "vbaPowerQuery"
'Dim apiSh As Worksheet
'
'Private Const apiUrlEntityRefType = "https://api.ravenpack.com/1.0/entity-reference?entity_type="
'Private Const apiUrlDatafile = "https://api.ravenpack.com/1.0/datafile"
'Private Const apiUrlJobs = "https://api.ravenpack.com/1.0/jobs"
'
'Option Explicit
'
'Sub Clear_Query_Tables()
'    Dim qt As QueryTable
'    Dim cn As Variant, qr As Variant
'
'    For Each cn In ActiveWorkbook.Connections
'        cn.Delete
'    Next
'
'    For Each qr In ActiveWorkbook.Queries
'        If qr.name = "Ravenpack_Data" Then
'            qr.Delete
'        End If
'    Next qr
'End Sub
'
'Sub Clear_Existing_Data()
'    Dim rng As Range: Set rng = ActiveWorkbook.ActiveSheet.UsedRange
'
'    With ActiveWindow
'        .SplitColumn = 0
'        .SplitRow = 0
'    End With
'
'    ActiveWindow.FreezePanes = False
'
'    ActiveSheet.Cells.Clear
'
'    Call clearShapes
'    Range("A:Z").ColumnWidth = 8.43
'    Cells(1, 1).Select
'
'End Sub
'
'
'Sub Query_Results(fileType As String)
'    Dim qt As QueryTable
'    Dim qry As WorkbookQuery
'    Dim wbkConn As WorkbookConnection
'    Dim M As String
'    Dim list_Obj As ListObject
'    Dim rng As Range
'
'    Set apiSh = ActiveWorkbook.Sheets(apiName)
'
'    Code_Run True
'
'    Clear_Query_Tables
'
'    ' Check if public API_KEY is empty
'    If api_key = "" Then
'        save_api_key
'        api_key = apiSh.Cells(1, 1)
'        If api_key = "" Then
'            Exit Sub
'        End If
'    End If
'
'    Set rng = ActiveSheet.UsedRange
'    'Ask for permission to delete existing data
'    If WorksheetFunction.CountA(rng) > 0 Then
'        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
'                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
'                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
'    End If
'
'    Clear_Existing_Data
'
'    'Create query to pull data
'
'    M = "let" & vbCrLf & _
'    "Source = Csv.Document(Web.Contents(" & Chr(34) & apiUrlEntityRefType & fileType & Chr(34) & _
'    ", [Headers=[API_KEY=" & Chr(34) & api_key & Chr(34) & "]]),[Delimiter=" & Chr(34) & "," & Chr(34) & ", Columns=6, Encoding=1252, QuoteStyle=QuoteStyle.None])" & vbCrLf & _
'    "in" & vbCrLf & _
'    "    Source"
'
'    StatusForm.status.Caption = "Loading Reference Data..."
'    StatusForm.Repaint
'    StatusForm.Show
'
'    Set qry = ActiveWorkbook.Queries.Add("Ravenpack_Data", M, "Data Pull")
'
'    Output_Query_Results qry, ActiveSheet
'
'    DoEvents
'
'    Clear_Query_Tables
'
'    'Clear table and remove formatting
'    For Each list_Obj In ActiveSheet.ListObjects
'        list_Obj.Unlist
'    Next
'
'    Set rng = ActiveSheet.UsedRange
'
'    With rng
'        .Interior.Color = xlNone
'    End With
'
'    ActiveSheet.Rows(1).EntireRow.Delete
'
'    StatusForm.status.Caption = vbNullString
'    StatusForm.Hide
'
'    Code_Run False
'End Sub
'
'Sub Test_Data_File()
'
'    Code_Run True
'
'    Query_Data_File Sheet5.Cells(2, 11), Sheet5.Cells(2, 12), Sheet5.Cells(3, 11), Sheet5.Cells(3, 12), _
'    Sheet5.Cells(5, 11), Sheet5.Cells(6, 11)
'
'    Code_Run False
'End Sub
'
'Sub Query_Data_File(ByVal start_date As String, ByVal start_date_time As String, _
'                    ByVal end_date As String, ByVal end_date_time As String, _
'                    ByVal dataset_uuid As String, ByVal api_key As String)
'
'    Dim fmt_start As String, fmt_end As String, fileURL As String, errMsg As String, M As String, token_id As String
'    Dim client As New WebClient, client2 As New WebClient
'    Dim request As New WebRequest, request2 As New WebRequest
'    Dim response As WebResponse, response2 As WebResponse
'    Dim requestBody As New Dictionary
'    Dim qry As WorkbookQuery
'    Dim list_Obj As ListObject
'    Dim rng As Range
'    Dim dataSh As Worksheet
'
'    Set dataSh = ActiveWorkbook.Sheets(ActiveSheet.name)
'
'    Clear_Query_Tables
'
'    DataRequestForm.status_label.Caption = "Requesting datafile..."
'
'    fmt_start = Format(CDate(start_date) + CDate(start_date_time), "YYYY-MM-DD hh:mm:ss")
'    fmt_end = Format(CDate(end_date) + CDate(end_date_time), "YYYY-MM-DD hh:mm:ss")
'
'    'Make initial request to get token
'    client.BaseUrl = apiUrlDatafile & "/" & dataset_uuid
'    request.Method = WebMethod.HttpPost
'    request.AddHeader "API_KEY", api_key
'    request.RequestFormat = WebFormat.JSON
'
'    requestBody.Add "start_date", fmt_start
'    requestBody.Add "end_date", fmt_end
'    requestBody.Add "format", "csv"
'    requestBody.Add "time_zone", "UTC"
'    requestBody.Add "compressed", False
'
'    Set request.Body = requestBody
'    Set response = client.Execute(request)
'
'    ' Print error message and exit
'    If response.StatusCode <> WebStatusCode.OK Then
'        errMsg = Response_Error_Handle(response)
'        MsgBox errMsg
'        Exit Sub
'    Else
'
'        token_id = response.Data("token")
'        client.BaseUrl = apiUrlJobs & "/" & token_id
'        request.Method = WebMethod.HttpGet
'        DataRequestForm.status_label.Caption = "Waiting for server to generate datafile..."
'
'        'CHECK TOKEN STATUS
'        Do
'            Application.Wait (Now + TimeValue("0:00:05")) 'wait 5 sec
'
'            ' Request the status of the job
'            Set response = client.Execute(request)
'            ' Print error message and exit
'            If response.StatusCode <> WebStatusCode.OK Then
'                DataRequestForm.status_label.Caption = "Error: " & response.Data("errors")
'                Exit Sub
'            ElseIf response.Data("status") = "completed" Then
'                DataRequestForm.status_label.Caption = "File Downloading..."
'                fileURL = response.Data("url")
'            End If
'
'        Loop While response.Data("status") <> "completed"
'        ' Datafile is now ready
'        DataRequestForm.status_label.Caption = "Datafile ready, downloading..."
'
'        If fileURL <> "" Then
'            DataRequestForm.status_label.Caption = "Loading Data..."
'
'            Clear_Existing_Data
'
'            M = "let" & vbCrLf & _
'            "Source = Csv.Document(Web.Contents(" & Chr(34) & fileURL & Chr(34) & _
'            ", [Headers=[API_KEY=" & Chr(34) & api_key & Chr(34) & "]]),[Delimiter=" & Chr(34) & "," & Chr(34) & ", Columns=6, Encoding=1252, QuoteStyle=QuoteStyle.None])" & vbCrLf & _
'            "in" & vbCrLf & _
'            "    Source"
'
'            Set qry = ActiveWorkbook.Queries.Add("Ravenpack_Data", M, "Data Pull")
'
'            Output_Query_Results qry, ActiveSheet
'
'            DoEvents
'
'            Clear_Query_Tables
'
'            'Clear table and remove formatting
'            For Each list_Obj In ActiveSheet.ListObjects
'                list_Obj.Unlist
'            Next
'
'            dataSh.Rows("1:1").Delete shift:=xlUp
'
'            Set rng = dataSh.UsedRange
'
'            With rng
'                .Interior.Color = xlNone
'            End With
'
'            With dataSh.Rows(1)
'                .Cells.HorizontalAlignment = xlLeft
'            End With
'
'            'FREEZE TOP ROW
'            With ActiveWindow
'                .SplitColumn = 0
'                .SplitRow = 1
'            End With
'
'            dataSh.Range("A1:AX1").Cells.Font.Bold = True
'            ActiveWindow.FreezePanes = True
'            ' Set focus on cell 1x1
'            Cells(1, 1).Select
'            dataSh.UsedRange.Columns.AutoFit
'            DataRequestForm.status_label.Caption = vbNullString
'            DataRequestForm.Hide
'        End If
'    End If
'
'End Sub
'
''Function to output Power Query Results
'Sub Output_Query_Results(query As WorkbookQuery, currentSheet As Worksheet)
'
'    With currentSheet.ListObjects.Add(SourceType:=0, Source:= _
'        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & query.name _
'        , Destination:=Range("$A$1")).QueryTable
'        .CommandType = xlCmdDefault
'        .CommandText = Array("SELECT * FROM [" & query.name & "]")
'        .RowNumbers = False
'        .FillAdjacentFormulas = False
'        .PreserveFormatting = True
'        .RefreshOnFileOpen = False
'        .BackgroundQuery = False
'        .RefreshStyle = xlInsertDeleteCells
'        .SavePassword = False
'        .SaveData = True
'        .AdjustColumnWidth = True
'        .RefreshPeriod = 0
'        .PreserveColumnInfo = False
'        .Refresh BackgroundQuery:=False
'    End With
'
'End Sub
'
'
