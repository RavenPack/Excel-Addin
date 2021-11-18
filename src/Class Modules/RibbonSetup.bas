Attribute VB_Name = "RibbonSetup"
Private mCurrentItemID As Variant
Dim prodType As String


Sub GetVisible(control As IRibbonControl, ByRef MakeVisible)
'PURPOSE: Show/Hide buttons based on how many you need (False = Hide/True = Show)

Select Case control.id
  Case "GroupA": MakeVisible = True
  Case "aButton01": MakeVisible = True
  Case "aButton02": MakeVisible = True
  
  Case "GroupB": MakeVisible = True
  Case "bButton01": MakeVisible = True
  Case "bButton03": MakeVisible = True
  
  Case "GroupC": MakeVisible = True
  Case "cButton01": MakeVisible = True
  
  Case "GroupE": MakeVisible = True
  Case "eButton01": MakeVisible = True
  
  Case "GroupF": MakeVisible = True
  Case "fButton01": MakeVisible = True
  
  Case "GroupG": MakeVisible = True
  Case "dropDown": MakeVisible = True
End Select

End Sub

Sub GetLabel(ByVal control As IRibbonControl, ByRef Labeling)
'PURPOSE: Determine the text to go along with your Tab, Groups, and Buttons

Select Case control.id
  
  Case "CustomTab": Labeling = "RavenPack"
  
  Case "GroupA": Labeling = ""
  Case "aButton01": Labeling = "Server Status"
  Case "aButton02": Labeling = "Set API KEY"
  
  Case "GroupB": Labeling = ""
  Case "bButton01": Labeling = "List Datasets"
  Case "bButton03": Labeling = "Data Request"
  
  Case "GroupC": Labeling = ""
  Case "cButton01": Labeling = "Map Entities"
  
  Case "GroupE": Labeling = ""
  Case "eButton01": Labeling = "Event Taxonomy"
  
  Case "GroupF": Labeling = ""
  Case "fButton01": Labeling = "Function Library"
  
End Select
   
End Sub

Sub GetImage(control As IRibbonControl, ByRef RibbonImage)
'PURPOSE: Tell each button which image to load from the Microsoft Icon Library
'TIPS: Image names are case sensitive, if image does not appear in ribbon after re-starting Excel, the image name is incorrect

Select Case control.id
  
  Case "aButton01": RibbonImage = "DatabaseCopyDatabaseFile"
  Case "aButton02": RibbonImage = "AdpPrimaryKey"
  
  Case "bButton01": RibbonImage = "ControlLayoutStacked"
  Case "bButton02": RibbonImage = "ControlLayoutTabular"
  Case "bButton03": RibbonImage = "ControlLayoutTabular"
  
  Case "cButton01": RibbonImage = "DiagramCycleInsertClassic"
  
  Case "eButton01": RibbonImage = "AccessFormDatasheet"
  
  Case "fButton01": RibbonImage = "Help"

End Select

End Sub

Sub GetSize(control As IRibbonControl, ByRef Size)
'PURPOSE: Determine if the button size is large or small

Const Large As Integer = 1
Const Small As Integer = 0

Select Case control.id
    
  Case "aButton01": Size = Large
  Case "aButton02": Size = Large
  
  Case "bButton01": Size = Large
  Case "bButton03": Size = Large
  
  Case "cButton01": Size = Large
  
  Case "eButton01": Size = Large
  
  Case "fButton01": Size = Large
  
End Select

End Sub

Sub RunMacro(control As IRibbonControl)
'PURPOSE: Tell each button which macro subroutine to run when clicked

Select Case control.id
  
  Case "aButton01": Application.Run "Button_Manager", "check_server_status"
  Case "aButton02": Application.Run "Button_Manager", "set_api_key"
  
  Case "bButton01": Application.Run "Button_Manager", "list_datasets"
  Case "bButton02": Application.Run "Button_Manager", "delete_datasets"
  Case "bButton03": Application.Run "Button_Manager", "data_request_button"
  
  Case "cButton01": Application.Run "Button_Manager", "entity_mapping_list_sub"
  
  Case "dButton01": Application.Run "Button_Manager", "cmdtReferenceFile"
  Case "dButton02": Application.Run "Button_Manager", "compReferenceFile"
  Case "dButton03": Application.Run "Button_Manager", "currReferenceFile"
  Case "dButton04": Application.Run "Button_Manager", "natlReferenceFile"
  Case "dButton05": Application.Run "Button_Manager", "orgaReferenceFile"
  Case "dButton06": Application.Run "Button_Manager", "peopReferenceFile"
  Case "dButton07": Application.Run "Button_Manager", "plceReferenceFile"
  Case "dButton08": Application.Run "Button_Manager", "prodReferenceFile"
  Case "dButton09": Application.Run "Button_Manager", "srceReferenceFile"
  
  Case "eButton01": Application.Run "Button_Manager", "taxonomy"
  
  Case "fButton01": Application.Run "Button_Manager", "FunctionLibraryForm_button"
  
 End Select
    
End Sub

Sub GetScreentip(control As IRibbonControl, ByRef Screentip)
'PURPOSE: Display a specific macro description when the mouse hovers over a button

Select Case control.id
  
  Case "aButton01": Screentip = "Check RavenPack server Status"
  Case "aButton02": Screentip = "Insert your API_KEY"
  
  Case "bButton01": Screentip = "List all of your datasets"
  Case "bButton03": Screentip = "Retrieve data for a dataset"
  
  Case "cButton01": Screentip = "Map a entities"
  
  Case "dButton01": Screentip = "Retrieve commodities reference data"
  Case "dButton02": Screentip = "Retrieve companies reference data"
  Case "dButton03": Screentip = "Retrieve currencies reference data"
  Case "dButton04": Screentip = "Retrieve nationalities reference data"
  Case "dButton05": Screentip = "Retrieve organizations reference data"
  Case "dButton06": Screentip = "Retrieve people reference data"
  Case "dButton07": Screentip = "Retrieve places reference data"
  Case "dButton08": Screentip = "Retrieve products reference data"
  Case "dButton09": Screentip = "Retrieve sources reference data"

  Case "eButton01": Screentip = "Retrieve RavenPack's Event Taxonomy File"
  
  Case "fButton01": Screentip = "Help for RavenPack functions"
  
  Case "dropDown": Screentip = "Select Ravenpack Product Type"
  
End Select

End Sub

Sub GetSelectedItemID(control As IRibbonControl, ByRef itemID As Variant)
    If IsEmpty(mCurrentItemID) Then
        mCurrentItemID = "item1"
    End If
    itemID = mCurrentItemID
    prodType = "item1"
End Sub


'' Definition of GetTheSelectedItemInDropDown which gets the selected item of the dropdown
' Sub GetTheSelectedItemInDropDown(control As IRibbonControl, id As String, index As Integer)
'    If control.id = "dropDown" Then
'        prodType = id
'    End If
' End Sub
'
'

'Definition for SetTheSelectedItemInDropDown which sets the value in the dropdown from the global variable
' Sub GetSelectedItemID(ByVal control As IRibbonControl, ByRef itemID As Variant)
'    If control.id = "dropDown" Then
'        itemID = "item1"
'        prodType = "item1"
'    End If
' End Sub

