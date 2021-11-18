Attribute VB_Name = "RavenPackFunctions"
Option Explicit

Public prodType As String

Private Const apiUrlJSON = "https://api.ravenpack.com/1.0/json"
Private Const apiUrlMapping = "https://api.ravenpack.com/1.0/entity-mapping"
Private Const apiUrlDatafile = "https://api.ravenpack.com/1.0/datafile"
Private Const apiUrlStatus = "https://api.ravenpack.com/1.0/status"
Private Const apiUrlBase = "https://api.ravenpack.com/1.0"

Private Const apiUrlJSONEdge = "https://api-edge.ravenpack.com/1.0/json"
Private Const apiUrlMappingEdge = "https://api-edge.ravenpack.com/1.0/entity-mapping"
Private Const apiUrlDatafileEdge = "https://api-edge.ravenpack.com/1.0/datafile"
Private Const apiUrlStatusEdge = "https://api-edge.ravenpack.com/1.0/status"
Private Const apiUrlBaseEdge = "https://api-edge.ravenpack.com/1.0"

Sub GetTheSelectedItemInDropDown(control As IRibbonControl, id As String, index As Integer)
   If control.id = "dropDown" Then
       prodType = id
   End If
End Sub

Function Set_Api_Source(api_source As String) As String

    If api_source = "Url_Map" Then
        
        If prodType = "item1" Then
            Set_Api_Source = apiUrlMapping
        Else
            Set_Api_Source = apiUrlMappingEdge
        End If
        
    ElseIf api_source = "Url_JSON" Then
    
        If prodType = "item1" Then
            Set_Api_Source = apiUrlJSON
        Else
            Set_Api_Source = apiUrlJSONEdge
        End If
        
    ElseIf api_source = "Url_Data" Then
        
        If prodType = "item1" Then
            Set_Api_Source = apiUrlDatafile
        Else
            Set_Api_Source = apiUrlDatafileEdge
        End If
        
    ElseIf api_source = "Status" Then
        
        If prodType = "item1" Then
            Set_Api_Source = apiUrlStatus
        Else
            Set_Api_Source = apiUrlStatusEdge
        End If
        
    ElseIf api_source = "Base" Then
        
        If prodType = "item1" Then
            Set_Api_Source = apiUrlBase
        Else
            Set_Api_Source = apiUrlBaseEdge
        End If
        
    End If
    
    
End Function

Public Function RPGetRecordCount(api_key As String, dataset_uuid As String, _
                                 start_date As String, end_date As String, _
                                 Optional time_zone As Variant)
                                 
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary
    Dim var As Variant
    Dim errMsg As String
    
    On Error GoTo ErrorHandle

    
    client.BaseUrl = Set_Api_Source("Url_Data") & "/" & dataset_uuid & "/count"
    'WebHelpers.EnableLogging = True
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    request.RequestFormat = WebFormat.JSON
    
    If IsMissing(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
    
    requestBody.Add "time_zone", CStr(time_zone)
    requestBody.Add "start_date", Format(start_date, "YYYY-MM-DD hh:mm:ss")
    requestBody.Add "end_date", Format(end_date, "YYYY-MM-DD hh:mm:ss")
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Check if we got the right response
    If response.StatusCode <> WebStatusCode.OK Then
        RPGetRecordCount = Null
        
        errMsg = Response_Error_Handle(response, Application.Caller)
        RPGetRecordCount = errMsg
        
        Exit Function
    Else
        RPGetRecordCount = response.Data("count")
    End If
    
    Exit Function
    
ErrorHandle:
    If Err.Number <> -2147210493 Then
        'ErrorHandling "RPGetRecordCount", ""
        RPGetRecordCount = Null
        Err.Clear
    Else
        Err.Clear
    End If
End Function

Public Function RPGetDailyValue(api_key As String, dataset_uuid As String, _
                                rp_entity_id As String, field_name As String, _
                                timestamp_utc As Date, Optional ByVal time_zone As Variant)
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary, conditions As New Dictionary, innerFn As New Dictionary
    Dim wrapperFn As New Dictionary, fn As New Dictionary, filters As New Dictionary
    Dim start_date As String, end_date As String, errMsg As String, fnName As String
    Dim blCust As Boolean
    
    On Error GoTo ErrorHandle
        
    client.BaseUrl = Set_Api_Source("Url_JSON") & "/" & dataset_uuid
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    request.RequestFormat = WebFormat.JSON
    conditions.Add "rp_entity_id", rp_entity_id
    
    If rp_entity_id <> "ROLLUP" Then
        filters.Add "rp_entity_id", rp_entity_id
    End If
    
    start_date = Format(timestamp_utc - 1, "yyyy-mm-dd") & " " & Format(timestamp_utc, "hh:mm:ss")
    end_date = Format(timestamp_utc, "yyyy-mm-dd") & " " & Format(timestamp_utc, "hh:mm:ss")
    
    If IsMissing(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
    
    blCust = False
    
    
    If LCase(field_name) = "sentiment" Then
        fnName = LCase(field_name)
        
        If prodType = "item1" Then
        
            innerFn.Add "field", "EVENT_SENTIMENT_SCORE"
            
        Else
        
            innerFn.Add "field", "EVENT_SENTIMENT"
            
        End If
        
        innerFn.Add "lookback", 91
        wrapperFn.Add "strength", innerFn
        fn.Add fnName, wrapperFn
        blCust = True
    ElseIf LCase(field_name) = "buzz" Then
        fnName = "buzz"
        innerFn.Add "field", "RP_ENTITY_ID"
        innerFn.Add "lookback", 91
        wrapperFn.Add "buzz", innerFn
        fn.Add fnName, wrapperFn
        blCust = True
    ElseIf LCase(field_name) = "volume" Then
        fnName = "volume"
        
        If prodType = "item1" Then
        
            innerFn.Add "field", "RP_STORY_ID"
        
        Else
        
            innerFn.Add "field", "RP_DOCUMENT_ID"
            
        End If
        
        wrapperFn.Add "cardinality", innerFn
        fn.Add fnName, wrapperFn
        blCust = True
'    ElseIf LCase(field_name) = "stddev" Then
'        fnName = "stddev"
'        innerFn.Add "field", "event_sentiment_score"
'        wrapperFn.Add "stddev", innerFn
'        fn.Add fnName, wrapperFn
'    ElseIf LCase(field_name) = "avg" Then
'        fnName = "avg"
'        innerFn.Add "field", "event_sentiment_score"
'        innerFn.Add "lookback", 1
'        wrapperFn.Add "avg", innerFn
'        fn.Add fnName, wrapperFn
    End If
    
    
    requestBody.Add "frequency", "daily"
    requestBody.Add "filters", filters
    
    If blCust Then
        requestBody.Add "custom_fields", Array(fn)
    End If
    
    requestBody.Add "fields", Array(LCase(field_name))
    requestBody.Add "conditions", conditions
    requestBody.Add "time_zone", time_zone
    requestBody.Add "start_date", start_date
    requestBody.Add "end_date", end_date
    
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Check if we got the right response
    If response.StatusCode <> WebStatusCode.OK Then
        
        errMsg = Response_Error_Handle(response, Application.Caller)
        'Error_Message errMsg
        
        RPGetDailyValue = errMsg
        
        'If VarType(Application.Caller) = 5 Then
        'Toggle_Change False
        'End If
        
        Exit Function
    Else
        RPGetDailyValue = response.Data("records")(1)(LCase(field_name))
    End If
    
    Exit Function
    
ErrorHandle:
    If Err.Number <> -2147210493 Then
        RPGetDailyValue = Err.Description
        Err.Clear
    Else
        Err.Clear
    End If
End Function

Public Function RPEntityName(api_key As String, rp_entity_id As String)
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary, identifiers As New Dictionary
    Dim errMsg As String
    
    On Error GoTo ErrorHandle
    
    If rp_entity_id = vbNullString Or Len(rp_entity_id) <> 6 Then
        RPEntityName = "Please enter a valid RP Entity ID"
        Exit Function
    End If
    
    client.BaseUrl = Set_Api_Source("Url_Map")

    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    identifiers.Add "name", rp_entity_id
    
    requestBody.Add "identifiers", Array(identifiers)
    request.RequestFormat = WebFormat.JSON
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Print error message and exit
    If response.StatusCode <> WebStatusCode.OK Or response.Data("identifiers_matched_count") <= 0 Then
        If response.Data("identifiers_matched_count") <= 0 Then
            errMsg = Response_Error_Handle(response, Application.Caller)
            
            If errMsg = vbNullString Then
                errMsg = "No matches returned for that RP Entity ID"
            End If
        Else
            errMsg = Response_Error_Handle(response, Application.Caller)
        End If
        
        RPEntityName = errMsg
    
    Else
        RPEntityName = response.Data("identifiers_mapped")(1)("rp_entities")(1)("rp_entity_name")
    End If
    
    Exit Function
    
ErrorHandle:
    If Err.Number <> -2147210493 Then
        RPEntityName = vbNullString
        Err.Clear
    Else
        Err.Clear
    End If
    
End Function

Public Function RPMapEntity(api_key As String, _
                            Optional ByVal entity_name As String, _
                            Optional ByVal entity_type As String, _
                            Optional ByVal ISIN As String, _
                            Optional ByVal CUSIP As String, _
                            Optional ByVal SEDOL As String, _
                            Optional ByVal listing As String, _
                            Optional ByVal matchDate As Date)
                            
    Dim errMsg As String
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary, identifiers As New Dictionary
                            
    ' Check that at least one optional arguement was supplied
    If IsMissing(entity_name) And IsMissing(entity_type) And _
        IsMissing(ISIN) And IsMissing(CUSIP) And IsMissing(SEDOL) And _
        IsMissing(listing) Then
        Exit Function
    End If
    
        
    client.BaseUrl = Set_Api_Source("Url_Map")
    
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    
    If Not IsMissing(entity_name) And entity_name <> vbNullString Then
        identifiers.Add "name", entity_name
    End If
    
    If Not IsMissing(entity_type) And entity_type <> vbNullString Then
        identifiers.Add "entity_type", entity_type
    End If
    
    If Not IsMissing(ISIN) And ISIN <> vbNullString Then
        identifiers.Add "isin", ISIN
    End If
    
    If Not IsMissing(CUSIP) And CUSIP <> vbNullString Then
        identifiers.Add "cusip", CUSIP
    End If
    
    If Not IsMissing(SEDOL) And SEDOL <> vbNullString Then
        identifiers.Add "sedol", SEDOL
    End If
    
    If Not IsMissing(listing) And listing <> vbNullString Then
        identifiers.Add "listing", listing
    End If
    
    If Not IsMissing(matchDate) Then
        identifiers.Add "date", Format(matchDate, "yyyy-mm-dd")
    End If
    
    requestBody.Add "identifiers", Array(identifiers)
    request.RequestFormat = WebFormat.JSON
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Print error message and exit
    If response.StatusCode <> WebStatusCode.OK Or response.Data("identifiers_matched_count") <= 0 Then
        If response.StatusCode <> WebStatusCode.OK Then
            errMsg = Response_Error_Handle(response, Nothing)
        Else
            errMsg = "No matches found."
        End If
        
        RPMapEntity = errMsg
    Else
        RPMapEntity = response.Data("identifiers_mapped")(1)("rp_entities")(1)("rp_entity_id")
    End If
End Function

'Updated
Private Function RPGetDailyEntityFn(api_key As String, rp_entity_id As String, _
                                    fnName As String, fn As Dictionary, _
                                    timestamp As Date, ByVal time_zone As Variant)
                                    
    Dim client As New WebClient
    Dim request As New WebRequest
    Dim response As WebResponse
    Dim requestBody As New Dictionary, filters As New Dictionary, conditions As New Dictionary
    Dim start_date As String, end_date As String, errMsg As String
    
    
    client.BaseUrl = Set_Api_Source("Url_JSON")
    request.Method = WebMethod.HttpPost
    request.AddHeader "API_KEY", api_key
    request.RequestFormat = WebFormat.JSON
    
    If rp_entity_id <> "ROLLUP" Then
        filters.Add "rp_entity_id", rp_entity_id
    End If
    
    conditions.Add "rp_entity_id", rp_entity_id
    start_date = Format(timestamp - 1, "yyyy-mm-dd") & " " & Format(timestamp, "hh:mm:ss")
    end_date = Format(timestamp, "yyyy-mm-dd") & " " & Format(timestamp, "hh:mm:ss")
    
    If IsMissing(time_zone) Or IsEmpty(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
    
    requestBody.Add "frequency", "daily"
    requestBody.Add "filters", filters
    requestBody.Add "custom_fields", Array(fn)
    requestBody.Add "fields", Array(fnName)
    requestBody.Add "conditions", conditions
    requestBody.Add "time_zone", CStr(time_zone)
    requestBody.Add "start_date", start_date
    requestBody.Add "end_date", end_date
    
    If prodType = "item1" Then
    
        requestBody.Add "product_version", "1.0"
        requestBody.Add "product", "rpa"
        
    Else
        
        requestBody.Add "product_version", "1.0"
        requestBody.Add "product", "edge"
        
    End If
    
    
    Set request.Body = requestBody
    client.timeOutMS = timeOutMilSec
    Set response = client.Execute(request)
    
    ' Check if we got the right response
    If response.StatusCode <> WebStatusCode.OK Then
        If IsArray(response.Data("records")) Then
            Debug.Print "Errors: " & response.Data("errors"), , "RavenPack"
            
            RPGetDailyEntityFn = Null
            Exit Function
        Else
            errMsg = Response_Error_Handle(response, Application.Caller)
        
            RPGetDailyEntityFn = errMsg
            Exit Function
        End If
    Else
        RPGetDailyEntityFn = response.Data("records")(1)(LCase(fnName))
    End If

End Function

'Updated
Public Function RPGetDailyEntitySentiment(api_key As Variant, rp_entity_id As Variant, _
                                          timestamp As Variant, Optional ByVal lookB As Variant, Optional ByVal time_zone As Variant)
    
    Dim innerFn As New Dictionary, wrapperFn As New Dictionary, fn As New Dictionary
    Dim fnName As String
    
    'Checks for valid date
    If Not Check_Date_Input(timestamp) Then
        RPGetDailyEntitySentiment = "Please enter a valid date."
        Exit Function
    End If
    
    If Len(rp_entity_id) <> 6 Then
        RPGetDailyEntitySentiment = "Please enter a valid RP Entity ID."
        Exit Function
    End If
    
    'Checks for valid lookback window
    If Not IsMissing(lookB) Then
        If IsError(lookB) Then
            RPGetDailyEntitySentiment = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
        
        If lookB <= 0 Then
            RPGetDailyEntitySentiment = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
    
        If IsError(Application.Evaluate("=" & CStr(lookB) & " * 1")) Then
            RPGetDailyEntitySentiment = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
    End If
    
    If IsMissing(time_zone) Or IsEmpty(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
    
    fnName = "sentiment"
    
    If prodType = "item1" Then
    
        innerFn.Add "field", "EVENT_SENTIMENT_SCORE"
        
    Else
    
        innerFn.Add "field", "EVENT_SENTIMENT"
        
    End If
    
    If IsMissing(lookB) Then
        innerFn.Add "lookback", 91
    Else
        innerFn.Add "lookback", CLng(lookB)
    End If
    
    wrapperFn.Add "strength", innerFn
    fn.Add fnName, wrapperFn
    
    RPGetDailyEntitySentiment = RPGetDailyEntityFn(CStr(api_key), CStr(rp_entity_id), fnName, fn, CDate(timestamp), CStr(time_zone))
End Function

'Updated
Public Function RPGetDailyEntityBuzz(api_key As Variant, rp_entity_id As Variant, _
                                          timestamp As Variant, Optional ByVal lookB As Variant, Optional ByVal time_zone As Variant)
                                          
    Dim innerFn As New Dictionary, wrapperFn As New Dictionary, fn As New Dictionary
    Dim fnName As String
    
    'Checks for valid date
    If Not Check_Date_Input(timestamp) Then
        RPGetDailyEntityBuzz = "Please enter a valid date."
        Exit Function
    End If
    
    If Len(rp_entity_id) <> 6 Then
        RPGetDailyEntityBuzz = "Please enter a valid RP Entity ID."
        Exit Function
    End If
    
    'Checks for valid lookback window
    If Not IsMissing(lookB) Then
        If IsError(lookB) Then
            RPGetDailyEntityBuzz = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
        
        If lookB <= 0 Then
            RPGetDailyEntityBuzz = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
        
        If IsError(Application.Evaluate("=" & CStr(lookB) & " * 1")) Then
            RPGetDailyEntityBuzz = "Please enter a valid integer for the lookback period."
            Exit Function
        End If
    End If
    
    If IsMissing(time_zone) Or IsEmpty(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
        
    fnName = "buzz"
    innerFn.Add "field", "RP_ENTITY_ID"
    
    If IsMissing(lookB) Then
        innerFn.Add "lookback", 91
    Else
        innerFn.Add "lookback", CLng(lookB)
    End If
    
    wrapperFn.Add "buzz", innerFn
    fn.Add fnName, wrapperFn
    
    RPGetDailyEntityBuzz = RPGetDailyEntityFn(CStr(api_key), CStr(rp_entity_id), fnName, fn, CDate(timestamp), CStr(time_zone))
End Function

'Updated and Working
Public Function RPGetDailyEntityVolume(api_key As Variant, rp_entity_id As Variant, _
                                          timestamp As Variant, Optional ByVal time_zone As Variant)
                                          
    Dim innerFn As New Dictionary, wrapperFn As New Dictionary, fn As New Dictionary
    Dim fnName As String
    
    'Checks for a valid length for the rp entity id
    If Len(rp_entity_id) <> 6 Then
        RPGetDailyEntityVolume = "Please enter a valid RP Entity ID."
        Exit Function
    End If
    
    'Checks for valid date
    If Not Check_Date_Input(timestamp) Then
        RPGetDailyEntityVolume = "Please enter a valid date."
        Exit Function
    End If
    
    If IsMissing(time_zone) Or IsEmpty(time_zone) Or CStr(time_zone) = vbNullString Then
        time_zone = "UTC"
    End If
    
    fnName = "volume"
    
    If prodType = "item1" Then
    
        innerFn.Add "field", "RP_STORY_ID"
        
    Else
    
        innerFn.Add "field", "RP_DOCUMENT_ID"
        
    End If
    
    wrapperFn.Add "cardinality", innerFn
    fn.Add fnName, wrapperFn
    
    RPGetDailyEntityVolume = RPGetDailyEntityFn(CStr(api_key), CStr(rp_entity_id), fnName, fn, CDate(timestamp), CStr(time_zone))
End Function





