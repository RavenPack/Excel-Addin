Attribute VB_Name = "RavenPackSubs"
Dim apiSh As Worksheet, actSh As Worksheet

Option Explicit

Public api_key As String
Private Const apiUrlEntityRefTest = "https://api.ravenpack.com/1.0/entity-reference/F9232E"
Private Const apiUrlEntityRefType = "https://api.ravenpack.com/1.0/entity-reference?entity_type="
Private Const apiUrlTaxonomy = "https://api.ravenpack.com/1.0/taxonomy"
Private Const apiUrlStatus = "https://api.ravenpack.com/1.0/status"
Private Const apiUrlDatasets = "https://api.ravenpack.com/1.0/datasets"
Private Const apiUrlDatafile = "https://api.ravenpack.com/1.0/datafile"
Private Const apiUrlJobs = "https://api.ravenpack.com/1.0/jobs"
Private Const apiUrlJSON = "https://api.ravenpack.com/1.0/json"
Private Const apiUrlMapping = "https://api.ravenpack.com/1.0/entity-mapping"

' Clear any table formating
Sub clearShapes()
    Dim shp As Shape
    
    For Each shp In ActiveSheet.Shapes
        If shp.Type = msoFormControl And shp.FormControlType = 1 Then
            shp.Delete
        End If
    Next
End Sub

'Clear sheet content
Sub clear_sheet()

    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 0
    End With
    
    ActiveWindow.FreezePanes = False
    rng.Clear
    clearShapes
    
    Range("A:Z").ColumnWidth = 8.43
    Cells(1, 1).Select
    

End Sub

' Get API_KEY
Sub set_api_key()
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim new_api_key As String
    Dim errMsg As String
    
    'Set Worksheet Variables
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    ' Ask the user for the API key
    new_api_key = Trim(InputBox("Please insert your RavenPack API_KEY", "RavenPack", api_key))
    
    ' If they didn't add one, then return an error
    If new_api_key = "" Then
        MsgBox "API Key must not be empty"
        Exit Sub
    End If
    
    ' Check if API_KEY is valid with an entity request F9232E
    
    client.BaseUrl = apiUrlEntityRefTest
    request.Method = WebMethod.HttpGet
    request.AddHeader "API_KEY", new_api_key
    
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    ' Check if we got the right response
    If response.StatusCode = WebStatusCode.GatewayTimeout Then
        errMsg = "Data Request Timed Out."

        MsgBox errMsg, , "RavenPack"
        Exit Sub
        
    ElseIf response.StatusCode <> WebStatusCode.OK Then
        errMsg = Response_Error_Handle(response, Nothing)
        MsgBox errMsg, , "RavenPack"
        Exit Sub
        
    Else
        api_key = new_api_key
        apiSh.Cells(1, 1).value = new_api_key
    End If
    
End Sub

' Get API_KEY
Sub save_api_key()

    Dim new_api_key As String
    
    'Set Worksheet Variables
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    If apiSh.Cells(1, 1).value = "" Then
        ' Ask the user for the API key
        new_api_key = Trim(InputBox("Please insert your RavenPack API_KEY", "RavenPack", api_key))
        
        ' If they didn't add one, then return an error
        If new_api_key = "" Then
            MsgBox "API Key must not be empty"
            Exit Sub
        End If
        
        ' Check if API_KEY is valid with an entity request F9232E
        Dim client As New WebClient, request As New WebRequest, response As WebResponse
        
        client.BaseUrl = apiUrlEntityRefTest
        request.Method = WebMethod.HttpGet
        request.AddHeader "API_KEY", new_api_key
        
        client.timeOutMS = timeOutMilSec
        Set response = client.Execute(request)
        ' Check if we got the right response
        If response.StatusCode = WebStatusCode.GatewayTimeout Then
            MsgBox "Timeout connecting to server"
        ElseIf response.StatusCode <> WebStatusCode.OK Then
            MsgBox "Errors: " & response.Data("errors"), , "RavenPack"
            Exit Sub
        Else
            api_key = new_api_key
            apiSh.Cells(1, 1).value = new_api_key
        End If
        
    Else
        new_api_key = apiSh.Cells(1, 1).value
    End If
    
End Sub

' Request data for a dataset
Sub data_request_button()
    'VARIABLE DATES
    Dim yesterday As String
    Dim today As String
    yesterday = Format(Date - 1, "yyyy-mm-dd")
    today = Format(Date, "yyyy-mm-dd")
    
    ' Set initial values on form
    'DataRequestForm.api_key_box.value = api_key
    DataRequestForm.start_date_box.value = yesterday
    DataRequestForm.end_date_box.value = today
    DataRequestForm.start_date_time_box.value = "00:00:00"
    DataRequestForm.end_date_time_box.value = "00:00:00"
    DataRequestForm.status_label.Caption = ""
    
    ' Center the form
    DataRequestForm.StartUpPosition = 0
    DataRequestForm.Left = Application.Left + (0.5 * Application.Width) - (0.5 * DataRequestForm.Width)
    DataRequestForm.Top = Application.Top + (0.5 * Application.Height) - (0.5 * DataRequestForm.Height)
    
    ' Show the form
    DataRequestForm.Show False
End Sub

Sub FunctionLibraryForm_button()
    FunctionLibraryForm.Show
End Sub


' Check RavenPack server's status
Sub check_server_status()
    Dim client As New WebClient
    Dim response As WebResponse
    
    ' Request the status from the server
    client.BaseUrl = apiUrlStatus
    Set response = client.GetJson("")
    If response.StatusCode = WebStatusCode.OK Then
        MsgBox "Status: " & response.Data("status"), , "RavenPack"
    Else
        'Debug.Print "Error: " & response.Content
        MsgBox "Bad response from server, http code: " & response.StatusCode
    End If
End Sub

' GET A LIST OF DATASETS
Sub list_datasets()
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim errMsg As String
    
    On Error GoTo ErrHandle
    
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    Code_Run True

    ' Check if public API_KEY is empty
    If api_key = "" Then
        save_api_key
        api_key = apiSh.Cells(1, 1)
        If api_key = "" Then
            Exit Sub
        End If
    End If
    
    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    
    'Ask for permission to delete existing data
    If WorksheetFunction.CountA(rng) > 0 Then
        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
    End If
    
    ' Send request of dataset list
    
    client.BaseUrl = apiUrlDatasets
    request.Method = WebMethod.HttpGet
    request.AddHeader "API_KEY", api_key
    StatusForm.status.Caption = "Fetching Dataset List..."
    StatusForm.Show
    
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Print error message and exit
    If response.StatusCode <> WebStatusCode.OK Then
        If response.StatusCode = 408 Then
            errMsg = "Error: Data Request Timed Out."
            MsgBox errMsg, , "RavenPack"
            
            Code_Run False
            Exit Sub
        End If
    
        errMsg = Response_Error_Handle(response, Nothing)
        MsgBox errMsg, , "RavenPack"
        
        Code_Run False
        Exit Sub
    Else
        ' Clear cell contents and switch off immediate updates to cells
        clear_sheet
        ' Set the header for the sheet
        Cells(1, 1).value = "DATASET_UUID"
        Cells(1, 2).value = "NAME"
        ' Loop through the results and put them on the sheet under the header
        Dim i As Integer
        For i = 2 To response.Data("datasets").Count + 1
            Cells(i, 2).value = response.Data("datasets")(i - 1)("name")
            Cells(i, 1).value = response.Data("datasets")(i - 1)("uuid")
        Next i
        
        'FREEZE TOP ROW
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        Range("A1:B1").Cells.Font.Bold = True
        ' Set the width of the columns
        Range("A:B").Columns.AutoFit
        ActiveWindow.FreezePanes = True
        ActiveSheet.Sort.SortFields.Clear
        
        With ActiveSheet.Sort
            .SortFields.Add key:=Range("B1"), Order:=xlAscending
            .SetRange Range("A:B")
            .header = xlYes
            .Apply
        End With
        ' Set focus on cell 1x1
        Cells(1, 1).Select
    End If
    
    StatusForm.Hide
    
    Code_Run False
    
    Exit Sub
    
ErrHandle:
    StatusForm.Hide
    Code_Run False
    ErrorHandling "List_Datasets", ""
End Sub
        
Sub getReferenceFile(entity_type)
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As New WebResponse
    Dim csvContent As String, location As String, csvLines() As String, csv As String, elements() As String, outputArr() As String
    Dim fileO As String, lineText As String, tempPath As String, errMsg As String
    Dim A As Variant, parsed As Variant
    Dim i As Long, j As Long, lines As Long, elCount As Long, start_row As Long
    Dim r As Range
    Dim csvColl As New Collection
    
    On Error GoTo ErrHandle
    
    Set apiSh = ActiveWorkbook.Sheets(apiName)
    Set actSh = ActiveSheet
    
    Code_Run True
    
    'Check if public API_KEY is empty
    If api_key = "" Then
        save_api_key
        api_key = apiSh.Cells(1, 1)
        If api_key = "" Then
            Exit Sub
        End If
    End If
    
    Dim rng As Range: Set rng = actSh.UsedRange
    'Ask for permission to delete existing data
    If WorksheetFunction.CountA(rng) > 0 Then
        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
    End If
    
    ' Make the HTTP request1
    
        
    'WebHelpers.EnableLogging = True
    client.BaseUrl = apiUrlEntityRefType & entity_type
    client.FollowRedirects = True
    request.Method = WebMethod.HttpGet
    request.AddHeader "API_KEY", api_key
    StatusForm.status.Caption = "Fetching Reference Data..."
    StatusForm.Show
    
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Print error message and exit
    
    If response.StatusCode = 303 Then
        Dim header As Dictionary
        For Each header In response.Headers
            If header("Key") = "Location" Then
                location = header("Value")
                Exit For
            End If
        Next
        
    ElseIf response.StatusCode <> WebStatusCode.OK Then
        If response.StatusCode = 408 Then
            errMsg = "Error: Data Request Timed Out."
            ActiveSheet.Cells(1, 1) = errMsg
            
            Code_Run False
            Exit Sub
        End If
        

        errMsg = Response_Error_Handle(response, Nothing)
        ActiveSheet.Cells(1, 1) = errMsg
        
        Code_Run False
        Exit Sub
        
    Else
    
        csvContent = response.Content
    End If
    
    If Not IsEmpty(location) And location <> "" Then
        client.BaseUrl = location
        client.timeOutMS = timeOutMilSec
        Set response = client.Execute(request)
        
        If response.StatusCode <> WebStatusCode.OK Then
            errMsg = Response_Error_Handle(response, Nothing)
            ActiveSheet.Cells(1, 1) = errMsg
            
            Exit Sub
        Else
            csvContent = response.Content
        End If
    End If
    
    If Not IsEmpty(csvContent) Then
        ' Clear current sheet
        Application.ScreenUpdating = False
        
        clear_sheet
        
        ' Set some specific formatting for reference files
        Range("A:H").ColumnWidth = 8.43
        Range("A:D").NumberFormat = "@"
        Range("E:F").NumberFormat = "yyyy-mm-dd"
        Range("E:F").HorizontalAlignment = xlRight
        Cells(1, 1).Select
        
        ' Load the CSV from the response content directly into sheet
        StatusForm.status.Caption = "Loading Reference Data..."
        StatusForm.Repaint
        
        csvLines = Split(response.Content, Chr(10))
        If Is_MAC Then
            elements = Split(csvLines(20), ",")
        Else
            elements = Split(csvLines(0), ",")
        End If

        elCount = UBound(elements)

        If Is_MAC Then
            start_row = 11
        Else
            start_row = 0
        End If

        
        ReDim outputArr(WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1)), UBound(elements))
        
        If UBound(csvLines) > 1048576 Then
            MsgBox "Maximum rows allowed by Excel exceeded. Truncating data imported."
            'Exit For
        End If

        
        
        For i = start_row To WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1))
            If i Mod 1000 = 0 Then
                StatusForm.status.Caption = "Loading " & i & " of " & CStr(WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1))) & " lines of Reference Data..."
                StatusForm.Repaint
                DoEvents
            End If
        
            If InStr(csvLines(i), ",") <> 0 Then
                j = 0
                Set csvColl = ParseCSVToCollection(csvLines(i))
                
                For Each A In csvColl(1)
                    outputArr(i, j) = A
                    j = j + 1
                Next
                

            End If
        Next i
        
        Set r = Range(Cells(1, 1), Cells(1, 1))
        r.Resize(UBound(outputArr, 1), UBound(outputArr, 2) + 1) = outputArr
            
        If Is_MAC Then
            ActiveSheet.Rows("1:11").EntireRow.Delete
        End If
        
        'FREEZE TOP ROW
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        Range("A1:F1").Cells.Font.Bold = True
        ActiveSheet.UsedRange.Columns.AutoFit
        ActiveWindow.FreezePanes = True
        ' Set focus on cell 1x1
        Cells(1, 1).Select
        Application.ScreenUpdating = True
    End If
    
    StatusForm.Hide
    
    Code_Run False
    
    Exit Sub

ErrHandle:
    StatusForm.Hide
    Code_Run False
    ErrorHandling "Get_Reference_File", Err.Description
End Sub

' COMMODITIES
Sub cmdtReferenceFile()
    getReferenceFile "CMDT"
End Sub

' COMPANIES
Sub compReferenceFile()
    getReferenceFile "COMP"
End Sub

' CURRENCIES
Sub currReferenceFile()
    getReferenceFile "CURR"
End Sub

' NATIONALITIES
Sub natlReferenceFile()
    getReferenceFile "NATL"
End Sub

' ORGANIZATIONS
Sub orgaReferenceFile()
    getReferenceFile "ORGA"
End Sub

' PEOPLE
Sub peopReferenceFile()
    getReferenceFile "PEOP"
End Sub

' PLACES
Sub plceReferenceFile()
    getReferenceFile "PLCE"
End Sub

' PRODUCTS
Sub prodReferenceFile()
    getReferenceFile "PROD"
End Sub

' SOURCES
Sub srceReferenceFile()
    getReferenceFile "SRCE"
End Sub
            
Sub taxonomy()
    Dim errMsg As String
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    
    Set apiSh = ActiveWorkbook.Sheets(apiName)

    'Check if public API_KEY is empty
    If api_key = "" Then
        save_api_key
        api_key = apiSh.Cells(1, 1)
        If api_key = "" Then
            Exit Sub
        End If
    End If
    
    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    
    'Ask for permission to delete existing data
    If WorksheetFunction.CountA(rng) > 0 Then
        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
    End If
    
    
    client.BaseUrl = apiUrlTaxonomy
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    StatusForm.status.Caption = "Fetching Taxonomy..."
    StatusForm.Show
    
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    ' Print error message and exit
    If response.StatusCode <> WebStatusCode.OK Then
        If response.StatusCode = 408 Then
            errMsg = "Error: Data Request Timed Out."
            ActiveSheet.Cells(1, 1) = errMsg
            ActiveSheet.Cells(1, 1).WrapText = False
            
            StatusForm.Hide
            
            Code_Run False
            Exit Sub
        End If
        
        errMsg = Response_Error_Handle(response, Nothing)
    
        Code_Run False
        ActiveSheet.Cells(1, 1) = errMsg
        ActiveSheet.Cells(1, 1).WrapText = False
        StatusForm.Hide
        Exit Sub
    Else
        ' Clear the active worksheet and switch off auto-updating
        Application.ScreenUpdating = False
        Call clear_sheet
        
        ' Set the headers by getting the first element and listing the keys
        ' Note, we're relying on the fact that the server sends the keys back
        ' in the correct order as headers
        Dim category_headers As Variant, i As Integer
        Set category_headers = response.Data("categories")(1)
        For i = 0 To category_headers.Count - 1
            Cells(1, i + 1).value = UCase(category_headers.keys()(i))
        Next i
        ' Start at row 2
        i = 2

        ' Loop through each entry
        Dim category As Variant
        For Each category In response.Data("categories")
            Dim k As Integer, fieldname As String
            For k = 0 To category.Count - 1
                fieldname = category.keys()(k)

                Dim value As Variant
                value = ""
                If fieldname = "valid_entity_types" Then
                    Dim p As Integer
                    For p = 0 To category.Item(fieldname).Count - 1
                        If value = "" Then
                            value = category.Item(fieldname)(p + 1)
                        Else
                            value = value & ", " & category.Item(fieldname)(p + 1)
                        End If
                    Next
                Else
                    value = category.Item(fieldname)
                End If
                
                Cells(i, k + 1).value = value
            Next k
                
            i = i + 1
        Next category
    End If
    'FREEZE TOP ROW
    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    Range("A1:J1").Cells.Font.Bold = True
    ' Set the width of the columns
    ActiveSheet.UsedRange.Columns.AutoFit
    ActiveWindow.FreezePanes = True
    
    ActiveSheet.Sort.SortFields.Clear
    With ActiveSheet.Sort
        .SortFields.Add key:=Range("A1"), Order:=xlAscending
        .SortFields.Add key:=Range("B1"), Order:=xlAscending
        .SortFields.Add key:=Range("C1"), Order:=xlAscending
        .SortFields.Add key:=Range("D1"), Order:=xlAscending
        .SortFields.Add key:=Range("E1"), Order:=xlAscending
        .SortFields.Add key:=Range("F1"), Order:=xlAscending
        .SetRange Range("A:J")
        .header = xlYes
        .Apply
    End With
    
    ' Set focus on cell 1x1
    Cells(1, 1).Select
    ' Switch back on screen updating
    Application.ScreenUpdating = True
    StatusForm.Hide
End Sub
                        
Sub entity_mapping_list_sub()
    Dim bGotRange As Boolean, bActivate As Boolean
    Dim rInput As Range, cell As Range
    Dim entity_list_col As New Collection, entity_sec_col As New Collection
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary, valueIndex As New Dictionary, dictTypes As New Dictionary, dictBody As New Dictionary, dictHeaders As New Dictionary
    Dim dictData As New Dictionary
    Dim rp_entity_id As String, rp_entity_name As String, rp_entity_type As String, data_type As String, errMsg As String, data_id_head As String
    Dim headStr As String, key As String
    Dim idx As Integer, i As Integer
    Dim rp_entity As Variant, requestData As Variant, var As Variant, matchedVar As Variant

    Set apiSh = ActiveWorkbook.Sheets(apiName)
    
    'On Error GoTo ErrHandle
    
    Code_Run True

    'Check if public API_KEY is empty
    If api_key = "" Then
        save_api_key
        api_key = apiSh.Cells(1, 1)
        If api_key = "" Then
            Exit Sub
        End If
    End If
    
    dictTypes.Add "rp_entity_id", "rp_entity_id"
    dictTypes.Add "entity_name", "name"
    dictTypes.Add "entity_type", "entity_type"
    dictTypes.Add "isin", "isin"
    dictTypes.Add "cusip", "cusip"
    dictTypes.Add "sedol", "sedol"
    dictTypes.Add "listing", "listing"
    'dictTypes.Add "matchdate", "matchDate"
    
'    'Ask for permission to delete existing data
'    Dim rng As Range: Set rng = ActiveSheet.UsedRange
'    If WorksheetFunction.CountA(rng) > 0 Then
'        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
'                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
'                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
'    End If
    
    'GET USER RANGE
    bActivate = False   ' True to re-activate the input range
    bGotRange = GetInputRange(rInput, "Please select a range of cells", _
                            "Entity Mapping", "", bActivate)
                            
    ' If nothing was selected then exit
    If Not bGotRange Then
        Exit Sub
    End If
    
    StatusForm.status.Caption = "Mapping Entities..."
    StatusForm.Show
    
    If rInput.Columns.Count = 1 Then
        ' Create a collection of entries
        For Each cell In rInput
            If cell.value <> "" Then
                entity_list_col.Add cell.value
            End If
        Next cell
    
    
        client.BaseUrl = apiUrlMapping
        request.Method = WebMethod.HttpPost
        request.AddHeader "API_KEY", api_key
        requestBody.Add "identifiers", entity_list_col
        request.RequestFormat = WebFormat.JSON
        Set request.Body = requestBody
        client.timeOutMS = timeOutMilSec
        Set response = client.Execute(request)
    
        ' Print error message and exit
        If response.StatusCode <> WebStatusCode.OK Then
            If response.StatusCode = 408 Then
                errMsg = "Error: Data Request Timed Out."
                MsgBox errMsg, , "RavenPack"
                
                Code_Run False
                Exit Sub
            End If
        
            errMsg = Response_Error_Handle(response, Nothing)
            MsgBox errMsg, , "RavenPack"
            GoTo exiS
        Else
            Application.ScreenUpdating = False
            ' Create dictionary for mapped values
            
            For i = 1 To response.Data("identifiers_mapped").Count
                requestData = response.Data("identifiers_mapped")(i)("request_data")
                ' Don't add duplicates
                If Not valueIndex.Exists(requestData) Then
                    valueIndex.Add requestData, i
                End If
            Next i
        
        
            For Each cell In rInput
                If cell.value <> "" Then
                    idx = valueIndex(cell.value)
                    
                    If idx <> 0 Then
                        If response.Data("identifiers_mapped")(idx)("rp_entities").Count > 0 Then
                            ' Always choose the first entity mapped
                            rp_entity_id = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_id")
                            rp_entity_name = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_name")
                            rp_entity_type = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_type")
                            ' Set the cell content with the returned information
                            cell.Offset(0, 1).NumberFormat = "@"
                            cell.Offset(0, 1).value = rp_entity_id
                            cell.Offset(0, 2).value = rp_entity_name
                            cell.Offset(0, 3).value = rp_entity_type
                        End If
                    End If
                End If
            Next cell
        End If
        
        Set request = Nothing
        
    ElseIf rInput.Columns.Count = 2 Then
        
        
        For Each cell In rInput
            Set dictBody = Nothing
            
            If cell.value <> "" And cell.Column = rInput.Column Then
                If Not dictBody.Exists(data_id_head) Then
                    dictBody.Add "name", cell.value
                End If

                If Not dictBody.Exists(data_type) Then
                    dictBody.Add "entity_type", ActiveSheet.Cells(cell.Row, cell.Column + 1).value
                End If
            
                entity_list_col.Add dictBody
            End If
        Next cell
        
        
        client.BaseUrl = apiUrlMapping
        request.Method = WebMethod.HttpPost
        request.AddHeader "API_KEY", api_key
        requestBody.Add "identifiers", entity_list_col
        request.RequestFormat = WebFormat.JSON
        Set request.Body = requestBody
        client.timeOutMS = timeOutMilSec
        Set response = client.Execute(request)
        
        ' Print error message and exit
        If response.StatusCode <> WebStatusCode.OK Then
            If response.StatusCode = 408 Then
                errMsg = "Error: Data Request Timed Out."
                MsgBox errMsg, , "RavenPack"
                
                Code_Run False
                Exit Sub
            End If
        
            errMsg = Response_Error_Handle(response, Nothing)
            MsgBox errMsg, , "RavenPack"
            GoTo exiS
        Else
            For i = 1 To response.Data("identifiers_mapped").Count
                requestData = response.Data("identifiers_mapped")(i)("request_data")("name")
                ' Don't add duplicates
                If Not valueIndex.Exists(requestData) Then
                    valueIndex.Add requestData, i
                End If
            Next i
        
        
            For Each cell In rInput
                If cell.Column = rInput.Column Then
                    If cell.value <> "" Then
                        idx = valueIndex(cell.value)
                        
                        If idx <> 0 Then
                            If response.Data("identifiers_mapped")(idx)("rp_entities").Count > 0 Then
                                Set matchedVar = Nothing
                            
                                For Each var In response.Data("identifiers_mapped")(idx)("rp_entities")
    
                                    If LCase(CStr(var.Item("rp_entity_type"))) = LCase(ActiveSheet.Cells(cell.Row, rInput.Column + 1).value) Then
                                        Set matchedVar = var
                                    End If
                                Next
                                
                                ' Always choose the first entity mapped
                                If Not matchedVar Is Nothing Then
                                    rp_entity_id = matchedVar.Item("rp_entity_id")
                                    rp_entity_name = matchedVar.Item("rp_entity_name")
                                    rp_entity_type = matchedVar.Item("rp_entity_type")
                                    ' Set the cell content with the returned information
                                    cell.Offset(0, 2).NumberFormat = "@"
                                    cell.Offset(0, 2).value = rp_entity_id
                                    cell.Offset(0, 3).value = rp_entity_name
                                    cell.Offset(0, 4).value = rp_entity_type
                                Else
                                    cell.Offset(0, 2).value = vbNullString
                                    cell.Offset(0, 3).value = vbNullString
                                    cell.Offset(0, 4).value = vbNullString
                                End If
                            End If
                        End If
                    End If
                End If
            Next cell
        End If
        
    ElseIf rInput.Columns.Count > 2 Then
    
        'Check to confrim headers conform to data types needed
        For i = 1 To rInput.Columns.Count
            headStr = LCase(WorksheetFunction.Substitute(Trim(ActiveSheet.Cells(rInput.Row, rInput.Column + (i - 1))), " ", "_"))
            
            If Not dictTypes.Exists(headStr) Then
                MsgBox "Data Type: " & CStr(ActiveSheet.Cells(rInput.Row, rInput.Column + (i - 1))) & " is invalid. Please input a correct data type into the header. Options include:" & _
                vbCrLf & "RP Entity ID, Entity Name, Entity Type, ISIN, CUSIP, SEDOL, and LISTING."
                
                GoTo exiS
            End If
            
            If Not dictHeaders.Exists(headStr) Then
                dictHeaders.Add headStr, rInput.Column + (i - 1)
            End If
        Next i

        
        'Populate body dictionary
        For Each cell In rInput
            Set dictBody = Nothing

            If cell.value <> "" And cell.Column = rInput.Column And cell.Row <> rInput.Row Then
                For Each var In dictHeaders.keys
                    If Not dictBody.Exists(var) Then
                        If ActiveSheet.Cells(cell.Row, dictHeaders.Item(var)).value <> vbNullString Then
                            dictBody.Add dictTypes(var), ActiveSheet.Cells(cell.Row, dictHeaders.Item(var))
                        End If
                    End If
                Next

                entity_list_col.Add dictBody
                Set dictBody = Nothing
            End If
        Next cell

        'Make URL call
        client.BaseUrl = apiUrlMapping
        request.Method = WebMethod.HttpPost
        request.AddHeader "API_KEY", api_key
        requestBody.Add "identifiers", entity_list_col
        request.RequestFormat = WebFormat.JSON
        Set request.Body = requestBody
        client.timeOutMS = timeOutMilSec
        Set response = client.Execute(request)

        ' Print error message and exit
        If response.StatusCode <> WebStatusCode.OK Then
            If response.StatusCode = 408 Then
                errMsg = "Error: Data Request Timed Out."
                MsgBox errMsg, , "RavenPack"
                
                Code_Run False
                Exit Sub
            End If
        
            errMsg = Response_Error_Handle(response, Nothing)
            MsgBox errMsg, , "RavenPack"
            GoTo exiS
        Else

            For i = 1 To response.Data("identifiers_mapped").Count
                For Each var In dictHeaders.keys
                    If key = "" Then
                        key = CStr(response.Data("identifiers_mapped")(i)("request_data")(dictTypes(var)))
                    Else
                        key = key & ":" & CStr(response.Data("identifiers_mapped")(i)("request_data")(dictTypes(var)))
                    End If
                Next
                    
                'requestData = response.Data("identifiers_mapped")(i)("request_data")
            
                ' Don't add duplicates
                If Not valueIndex.Exists(key) Then
                    valueIndex.Add key, i
                End If
                
                key = ""
                    
            Next i


            For Each cell In rInput
                If cell.Row <> rInput.Row And cell.Column = rInput.Column Then
                    If cell.value <> "" Then
                        'Construct unique key
                        For i = rInput.Column To rInput.Column + rInput.Columns.Count - 1
                            If i = rInput.Column Then
                                key = ActiveSheet.Cells(cell.Row, i)
                            Else
                                key = key & ":" & CStr(ActiveSheet.Cells(cell.Row, i))
                            End If
                        Next i
                        
                        If valueIndex.Exists(key) Then
                            idx = valueIndex.Item(key)
                            
                            If idx <> 0 Then
                                If response.Data("identifiers_mapped")(idx)("rp_entities").Count > 0 Then
                                    ' Always choose the first entity mapped
                                    rp_entity_id = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_id")
                                    rp_entity_name = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_name")
                                    rp_entity_type = response.Data("identifiers_mapped")(idx)("rp_entities")(1)("rp_entity_type")
        
                                    ' Set the cell content with the returned information
                                    cell.Offset(0, rInput.Columns.Count).NumberFormat = "@"
                                    cell.Offset(0, rInput.Columns.Count).value = rp_entity_id
                                    cell.Offset(0, rInput.Columns.Count + 1).value = rp_entity_name
                                    cell.Offset(0, rInput.Columns.Count + 2).value = rp_entity_type
                                End If
                            End If
                        End If
                    End If
                End If
            Next cell
        End If

        Set request = Nothing
    End If
    
    StatusForm.Hide
    Code_Run False
    
    Exit Sub
    
exiS:
    StatusForm.Hide
    Code_Run False
    Exit Sub
    
ErrHandle:
    StatusForm.Hide
    Code_Run False
    ErrorHandling "Entity Mapping", ""
End Sub

Function GetInputRange(rInput As Excel.Range, _
                       sPrompt As String, _
                       sTitle As String, _
                       Optional ByVal sDefault As String, _
                       Optional ByVal bActivate As Boolean, _
                       Optional X, Optional Y) As Boolean
' rInput:    The Input Range which returns to the caller procedure
' bActivate: If True user's input range will be re-activated
'
' The other arguments are standard InputBox arguments.
' sPrompt & sTitle should be supplied from the caller proccedure
' but sDefault will be completed below if empty
'
' GetInputRange returns True if rInput is successfully assigned to a Range

    Dim bGotRng As Boolean
    Dim bEvents As Boolean
    Dim nAttempt As Long
    Dim sAddr As String
    Dim vReturn

    On Error Resume Next
    
    If Len(sDefault) = 0 Then
        If TypeName(Application.Selection) = "Range" Then
            sDefault = "=" & Application.Selection.Address
            ' InputBox cannot handle address/formulas over 255
            If Len(sDefault) > 240 Then
                sDefault = "=" & Application.ActiveCell.Address
            End If
        ElseIf TypeName(Application.ActiveSheet) = "Chart" Then
            sDefault = " first select a Worksheet"
        Else
            sDefault = " Select Cell(s) or type address"
        End If
    End If
    
    Set rInput = Nothing    ' start with a clean slate
    For nAttempt = 1 To 3  ' give user 3 attempts for typos
        vReturn = False
        vReturn = Application.InputBox(sPrompt, sTitle, sDefault, X, Y, Type:=0)
        If False = vReturn Or Len(vReturn) = 0 Then
            Exit For    ' user cancelled
        Else
            sAddr = vReturn
            ' The address (or formula) could be in A1 or R1C1 style,
            ' w/out an "=" and w/out embracing quotes, depends if the user
            ' selected cells, typed an address, or accepted the default
            If Left$(sAddr, 1) = "=" Then sAddr = Mid$(sAddr, 2, 256)
            If Left$(sAddr, 1) = Chr(34) Then sAddr = Mid$(sAddr, 2, 255)
            If Right$(sAddr, 1) = Chr(34) Then sAddr = Left$(sAddr, Len(sAddr) - 1)
            '  will fail if R1C1 address
            Set rInput = Application.Range(sAddr)
            If rInput Is Nothing Then
                sAddr = Application.ConvertFormula(sAddr, xlR1C1, xlA1)
                Set rInput = Application.Range(sAddr)
                bGotRng = Not rInput Is Nothing
            Else
                bGotRng = True
            End If
        End If

        If bGotRng Then
            If bActivate Then   ' optionally re-activate the Input range
                On Error GoTo errH
                bEvents = Application.EnableEvents
                Application.EnableEvents = False
                If Not Application.ActiveWorkbook Is rInput.Parent.Parent Then
                    rInput.Parent.Parent.Activate    ' Workbook
                End If
                If Not Application.ActiveSheet Is rInput.Parent Then
                    rInput.Parent.Activate    ' Worksheet
                End If
                rInput.Select    ' Range
            End If
            Exit For
        ElseIf nAttempt < 3 Then
            ' Failed to get a valid range, maybe a typo
            If MsgBox("Invalid reference, do you want to try again ?", _
                      vbOKCancel, sTitle) <> vbOK Then
                Exit For
            End If
        End If
    Next    ' nAttempt
cleanUp:
    On Error Resume Next
    If bEvents Then
        Application.EnableEvents = True
    End If
    GetInputRange = bGotRng
    Exit Function
errH:
    Set rInput = Nothing
    bGotRng = False
    Resume cleanUp
End Function

Sub DatafileRequest(ByVal start_date As String, ByVal start_date_time As String, _
                    ByVal end_date As String, ByVal end_date_time As String, _
                    ByVal dataset_uuid As String, ByVal api_key As String)
    DataRequestForm.status_label.Caption = "Requesting datafile..."
    
    Dim client As New WebClient, client2 As New WebClient
    Dim request As New WebRequest, request2 As New WebRequest
    Dim response As WebResponse, response2 As WebResponse
    Dim requestBody As New Dictionary
    Dim fileURL As String, errMsg As String, token_id As String, csvLines() As String, elements() As String, outputArr() As String
    Dim i As Long, j As Long, elCount As Long, start_row As Long
    Dim r As Range
    Dim csvColl As New Collection
    Dim A As Variant
    
    On Error GoTo ErrHandle
    
    Set actSh = ActiveSheet
    
    Code_Run True
        
    client.BaseUrl = apiUrlDatafile & "/" & dataset_uuid
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    request.RequestFormat = WebFormat.JSON
    
    requestBody.Add "start_date", start_date & " " & start_date_time
    requestBody.Add "end_date", end_date & " " & end_date_time
    requestBody.Add "format", "csv"
    requestBody.Add "time_zone", "UTC"
    requestBody.Add "compressed", False
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    ' Print error message and exit
    If response.StatusCode <> WebStatusCode.OK Then
        If response.StatusCode = 408 Then
            errMsg = "Error: Data Request Timed Out."
            MsgBox errMsg, , "RavenPack"
            
            Code_Run False
            Exit Sub
        End If
    
        errMsg = Response_Error_Handle(response, Nothing)
        DataRequestForm.status_label.Caption = errMsg
        
        Code_Run False
        
        Exit Sub
        
    Else
        token_id = response.Data("token")
        client.BaseUrl = apiUrlJobs & "/" & token_id
        
        request.Method = WebMethod.HttpGet
        Set request.Body = New Dictionary
        
        DataRequestForm.status_label.Caption = "Waiting for server to generate datafile..."
        'CHECK TOKEN STATUS
        Do
            Application.Wait (Now + TimeValue("0:00:05")) 'wait 5 sec
            ' Request the status of the job
            client.timeOutMS = timeOutMilSec
            Set response = client.Execute(request)
            ' Print error message and exit
            If response.StatusCode <> WebStatusCode.OK Then
                DataRequestForm.status_label.Caption = "Error: " & response.Data("errors")
                Exit Sub
            ElseIf response.Data("status") = "completed" Then
                DataRequestForm.status_label.Caption = "File Downloading..."
                fileURL = response.Data("url")
            End If
        Loop While response.Data("status") <> "completed"
        
        
        ' Datafile is now ready
        DataRequestForm.status_label.Caption = "Datafile ready, downloading..."
        
        If fileURL <> "" Then
            client2.BaseUrl = fileURL
            client2.FollowRedirects = True
            
            request2.Method = WebMethod.HttpGet
            request2.AddHeader "API_KEY", api_key
            Set response2 = client2.Execute(request2)
            
            ' Print error message and exit
            If response2.StatusCode <> WebStatusCode.OK Then
                If response2.StatusCode = 408 Then
                    errMsg = "Error: Data Request Timed Out."
                    MsgBox errMsg, , "RavenPack"
                    
                    Code_Run False
                    Exit Sub
                End If
            
                errMsg = Response_Error_Handle(response2, Nothing)
                DataRequestForm.status_label.Caption = errMsg
                
                Code_Run False
                
                Exit Sub
            Else
                DataRequestForm.status_label.Caption = "Loading Data..."
                Application.ScreenUpdating = False
                ' Clear the active sheet
                clear_sheet
                ' Load the CSV from the response content directly into sheet
                
                
                csvLines = Split(response2.Content, Chr(10))
                If Is_MAC Then
                    elements = Split(csvLines(20), ",")
                Else
                    elements = Split(csvLines(0), ",")
                End If

                elCount = UBound(elements)

                If Is_MAC Then
                    start_row = 11
                Else
                    start_row = 0
                End If

                elCount = UBound(elements)
                
                ReDim outputArr(WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1)), UBound(elements))
                
                If UBound(csvLines) > 1048576 Then
                    DataRequestForm.status_label.Caption = "Maximum rows allowed by Excel exceeded. Truncating data imported."
                End If
                
                For i = 0 To WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1))
                    If i Mod 1000 = 0 Then
                        DataRequestForm.status_label.Caption = "Loading " & i & " of " & CStr(WorksheetFunction.Min(actSh.Rows.Count, UBound(csvLines, 1))) & " lines of Reference Data..."
                        StatusForm.Repaint
                        DoEvents
                    End If
                
                    If InStr(csvLines(i), ",") <> 0 Then
                        j = 0
                        Set csvColl = ParseCSVToCollection(csvLines(i))
                        
                        For Each A In csvColl(1)
                            outputArr(i, j) = A
                            j = j + 1
                        Next
                        
        
                    End If
                Next i
                
                Set r = Range(Cells(1, 1), Cells(1, 1))
                r.Resize(UBound(outputArr, 1), UBound(outputArr, 2) + 1) = outputArr
                
                If Is_MAC Then
                    ActiveSheet.Rows("1:11").EntireRow.Delete
                End If
                
                With ActiveSheet.Rows(1)
                    .Cells.HorizontalAlignment = xlLeft
                End With
                'FREEZE TOP ROW
                With ActiveWindow
                    .SplitColumn = 0
                    .SplitRow = 1
                End With
                Range("A1:AX1").Cells.Font.Bold = True
                ActiveWindow.FreezePanes = True
                ' Set focus on cell 1x1
                Cells(1, 1).Select
                ActiveSheet.UsedRange.Columns.AutoFit
                Application.ScreenUpdating = True
                DataRequestForm.Hide
            End If
        End If
    End If
    
    Code_Run False
    
    Exit Sub

ErrHandle:
    DataRequestForm.Hide
    Code_Run False
    ErrorHandling "Data_File_Request", ""
End Sub

Public Sub DataRequest(ByVal start_date As String, ByVal start_date_time As String, _
                       ByVal end_date As String, ByVal end_date_time As String, _
                       ByVal dataset_uuid As String, ByVal api_key As String)
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary
    Dim datasetFields As Variant
    Dim fileURL As String, start_date_fmt As String, end_date_fmt As String, errMsg As String
    
    Code_Run True
    
    
    
    'Ask for permission to delete existing data
    Dim rng As Range: Set rng = ActiveSheet.UsedRange
    
    If WorksheetFunction.CountA(rng) > 0 Then
        If MsgBox(Prompt:="Your current spreadsheet is not empty. The existing data will be purged " + _
                          "by the pending action with no recovery option. Are you sure you would like to proceed?", _
                  Buttons:=vbYesNo + vbExclamation, Title:="CAUTION!") = 7 Then End
    End If
    
    ' Check that the required parameters are non-null
    If start_date = "" Or start_date_time = "" Or end_date = "" Or _
        end_date_time = "" Or dataset_uuid = "" Or api_key = "" Or _
        dataset_uuid = "" Then
        MsgBox "Please complete all fields", , "RavenPack"
        Code_Run False
        
        Exit Sub
    End If
    
    DataRequestForm.status_label.Caption = "Requesting Data..."
    ' Fetch the dataset headers
    
    ' Set time parameters with the correct ANSI format
    start_date_fmt = Format(start_date, "yyyy-mm-dd") & " " & Format(start_date_time, "hh:mm:ss")
    end_date_fmt = Format(end_date, "yyyy-mm-dd") & " " & Format(end_date_time, "hh:mm:ss")
    
    If CDate(start_date_fmt) > CDate(end_date_fmt) Then
        MsgBox "End time must occur after the start time."
        
        Code_Run False
        Exit Sub
    End If
    
    ' reuse our http objects
    client.BaseUrl = apiUrlJSON & "/" & dataset_uuid
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    request.RequestFormat = WebFormat.JSON
    requestBody.Add "start_date", start_date_fmt
    requestBody.Add "end_date", end_date_fmt
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' If we exceeded the allowed limit then Ask the user if they would like to make a datafile request
    If response.StatusCode = 400 Then
        Dim recordCount As Long
        
        If IsNull(RPGetRecordCount(api_key, dataset_uuid, start_date_fmt, end_date_fmt, "UTC")) Then
            DataRequestForm.status_label.Caption = "No Data Returned"
            
            Code_Run False
            
            Exit Sub
        End If
        
        recordCount = RPGetRecordCount(api_key, dataset_uuid, start_date_fmt, end_date_fmt, "UTC")
        
        If recordCount > 1048576 Then
            DataRequestForm.status_label.Caption = "Number of records requested exceeds the maximum allowed by Excel in one sheet."
            Code_Run False
            Exit Sub
        End If
        
        ' Make a DataFile request
        Dim datafilep As Integer
        
        datafilep = MsgBox("The request exceeds the maximum rows via the JSON endpoint. " & _
                        "Would you like to generate a datafile?", vbOKCancel + vbDefaultButton1, "RavenPack")
        
        If datafilep = vbOK Then
            'If Is_MAC Then
            Call DatafileRequest(start_date, start_date_time, end_date, end_date_time, dataset_uuid, api_key)
'            Else
'                Query_Data_File start_date, start_date_time, end_date, end_date_time, dataset_uuid, api_key
'            End If
        Else
            ' Otherwise just exit
            DataRequestForm.status_label.Caption = "Request cancelled..."
            Code_Run False
            Exit Sub
        End If
    ' If any other status, print the error and exit
    ElseIf response.StatusCode <> WebStatusCode.OK Then
        If response.StatusCode = 408 Then
            DataRequestForm.status_label.Caption = "Error: Data Request Timed Out."
            Code_Run False
            Exit Sub
        End If
    
        errMsg = Response_Error_Handle(response, Nothing)
    
        DataRequestForm.status_label.Caption = "Error: " & errMsg
        Code_Run False
        Exit Sub
        
    ' Catch all other errors
    ElseIf response.Data("errors") Then
        errMsg = Response_Error_Handle(response, Nothing)
    
        DataRequestForm.status_label.Caption = "Error: " & errMsg
        Code_Run False
        Exit Sub
        
    ' Otherwise use the results from the JSON call
    Else
        ' Switch off cell auto-updating so updates are quick
        Application.ScreenUpdating = False
        ' Clear the active sheet
        Call clear_sheet
        
        ' Insert the results of the json query
        ' Start with the header, assuming the data comes in the correct order
        If IsEmpty(response.Data("records")) Or response.Data("records").Count = 0 Then
            DataRequestForm.status_label.Caption = "No Data..."
            Code_Run False
            
            Exit Sub
        Else
            DataRequestForm.status_label.Caption = "Loading Data..."
            DataRequestForm.Repaint
            Dim i As Integer, firstRecord As Dictionary, field As String
            ' Get header from first record
            Set firstRecord = response.Data("records")(1)
            For i = 0 To firstRecord.Count - 1
                Cells(1, i + 1).value = UCase(firstRecord.keys()(i))
            Next i
            ' Then insert the data in reverse order
            i = response.Data("records").Count + 1
            
            Dim Item As Object
            
            For Each Item In response.Data("records")
                DoEvents
                
                Dim w As Integer, fieldname As String
                For w = 0 To Item.Count - 1
                    'DoEvents
                    
                    fieldname = LCase(Item.keys()(w))
                    ' Format cells
                    If fieldname = "rp_entity_id" Or fieldname = "rp_source_id" Or fieldname = "related_entity" Then
                        ActiveSheet.Cells(i, w + 1).NumberFormat = "@"
                    ElseIf fieldname = "timestamp_utc" Then
                        ActiveSheet.Cells(i, w + 1).NumberFormat = "yyyy-mm-dd hh:mm:ss.000"
                    ElseIf fieldname = "event_start_date_utc" Or fieldname = "event_end_date_utc" Or _
                           fieldname = "reporting_period" Or fieldname = "reporting_start_date_utc" Or _
                           fieldname = "reporting_end_date_utc" Then
                        ActiveSheet.Cells(i, w + 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
                    End If
                    'ActiveSheet.Cells(i, w + 1).value = Item(fieldname)
                    ActiveSheet.Cells(i, w + 1).value = Item(Item.keys()(w))
                Next w
                i = i - 1
            Next Item
        End If
        With ActiveSheet.Rows(1)
            .Cells.HorizontalAlignment = xlLeft
        End With
        
        'FREEZE TOP ROW
        With ActiveWindow
            .SplitColumn = 0
            .SplitRow = 1
        End With
        Range("A1:AX1").Cells.Font.Bold = True
        ActiveSheet.UsedRange.Columns.AutoFit
        ActiveWindow.FreezePanes = True
        ' Set focus on cell 1x1
        Cells(1, 1).Select
        Application.ScreenUpdating = True
        DataRequestForm.Hide
    End If
    
    Code_Run False
End Sub
