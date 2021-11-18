VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    
    prodType = "item1"
    
    Verify_Perm
    
    On Error Resume Next
    
    Application.MacroOptions _
        macro:="RPGetDailyEntitySentiment", _
        Description:="Return the 91-day sentiment strength for an entity on a particular day", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "6 character RavenPack Entity ID", _
                                    "The datetime that represents the end of the day to calculate the sentiment", _
                                    "The lookback period in days desired. 91 by default.", _
                                    "The time zone that the datetime is in for the request. UTC by default")
    
    Application.MacroOptions _
        macro:="RPGetDailyEntityBuzz", _
        Description:="Return the average  media buzz for an entity on a particular day", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "6 character RavenPack Entity ID", _
                                    "The datetime that represents the end of the day to calculate the buzz", _
                                    "The lookback period in days desired. 91 by default.", _
                                    "The time zone that the datetime is in for the request. UTC by default")
    
    Application.MacroOptions _
        macro:="RPGetDailyEntityVolume", _
        Description:="Return the total volume of stories for an entity on a particular day", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "6 character RavenPack Entity ID", _
                                    "The datetime that represents the end of the day to calculate the volume", _
                                    "The time zone that the datetime is in for the request. UTC by default")
    
    Application.MacroOptions _
        macro:="RPGetDailyValue", _
        Description:="Return the value for a particular indicator for a daily dataset on a particular day", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "The dataset_uuid for the daily dataset", _
                                    "6 character RavenPack Entity ID", _
                                    "The field name of the indicator that should be returned", _
                                    "The datetime that represents the end of the day to calculate the field value", _
                                    "The time zone that the datetime is in for the request. UTC by default")
    
    Application.MacroOptions _
        macro:="RPEntityName", _
        Description:="Map to the company ID give the RP_ENTITY_ID", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "6 character RavenPack Entity ID")

    
    Application.MacroOptions _
        macro:="RPMapEntity", _
        Description:="Map to the RP_ENTITY_ID given a set of identifying information", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "The name of the entity", _
                                    "The entity_type associated with this entity, e.g. COMP", _
                                    "An ISIN associated with this entity", _
                                    "A CUSIP associated with this entity", _
                                    "A SEDOL associated with this entity", _
                                    "A Listing, composed of a Mic code and ticker associated with this entity, e.g. XNAS:MSFT", _
                                    "A datetime for which the identifying information is valid")

    Application.MacroOptions _
        macro:="RPGetRecordCount", _
        Description:="Return the count for a particular dataset on a particular day", _
        category:=5, _
        ArgumentDescriptions:=Array("Key supplied by RavenPack for accessing the Rest API", _
                                    "The dataset_uuid for the daily dataset", _
                                    "The start date and time in ANSI format", _
                                    "The end date and time in ANSI format", _
                                    "The time zone that the datetime is in for the request. UTC by default")
                                    
                    
    On Error GoTo 0

End Sub







