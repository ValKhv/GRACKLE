VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cGoogleCharts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
' Class creates html file with google charts settings
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
'******************************************************************************************************************************************************
Option Explicit

    Private Const MOD_NAME As String = "cGoogleCharts"
'**************************
    Public Enum TCHARTTYPE
         CHR_BAR = 0
         CHR_BUBBLE = 1
         CHR_COLUMNS = 2
         CHR_DONUTS = 3
         CHR_GANTT = 4
         CHR_GAUGE = 5
         CHR_GEO = 6
         CHR_LINE = 7
         CHR_PIE = 8
         CHR_TIMELINES = 9
    End Enum
    
    Private Type TTASK
        TaskID As String
        TaskName As String
        Resource As String
        StartDate As String
        EndDate As String
        Duration As String
        PercentComplete As String
        Dependencies As String
    End Type
    
    Private Type TConfig
          ChartName As String
          ChartID As String
          DrawFuncName As String
          ThePackageName As String
          TheClass As String
          fileName As String
          APIProvider As String
    End Type
    
'**************************
    Private m_Res As String              ' The string (html body) to write in html file
    Private TC As TConfig                ' The Config for Chart
    Private m_IsMaterial As Boolean      ' Bolean flag to use material library
    
    Private m_Folder As String           ' The folder
    
    Private m_Path As String             ' Full path to html file
    Private m_ChartType As TCHARTTYPE    ' The chart Type
    
    Private m_OPTIONS As cStack          ' Options in Charts
    Private m_sOPT As String             ' Options as string
    
    
    Private m_Width As Integer           ' Width for div or chart in ptx
    Private m_Height As Integer          ' Height for div or chart in ptx
    Private m_Title As String            ' Chart title
    
    Private m_ChartLib As String         ' Extension for Package (need for some charts)
    
    Private m_Data As cVTable            ' Array of data for charting
    
    Private m_JavaScript As String       ' Intrinsic JavaScript
    
    Private m_sEXT As String             ' Extension for Func Presentation
'**************************

Public Property Get IsMaterial() As Boolean
         IsMaterial = m_IsMaterial
End Property
Public Property Let IsMaterial(bMaterial As Boolean)
         m_IsMaterial = bMaterial
End Property

Public Property Get Folder() As String
    Folder = m_Folder
End Property
Public Property Let Folder(sFolder As String)
    m_Folder = sFolder
End Property


Public Property Get ChartTitle() As String
    ChartTitle = m_Title
End Property
Public Property Let ChartTitle(sTitle As String)
    m_Title = sTitle
End Property

Public Property Get height() As Integer
       height = m_Height
End Property
Public Property Let height(iHeight As Integer)
        m_Height = iHeight
End Property

Public Property Get Width() As Integer
        Width = m_Width
End Property
Public Property Let Width(IWidth As Integer)
        m_Width = IWidth
End Property




Public Property Get ChartType() As TCHARTTYPE
     ChartType = m_ChartType
End Property
Public Property Let ChartType(ChrtType As TCHARTTYPE)
     m_ChartType = ChrtType
     
     ' SET OTHER CHARACTERS
     Call CreateConfig(m_ChartType)
     
End Property

Public Property Get ChartID() As String
       ChartID = TC.ChartID
End Property
Public Property Let ChartID(sID As String)
       TC.ChartID = sID
End Property

Public Property Get Package() As String
    Package = TC.ThePackageName
End Property
Public Property Let Package(sPackageName As String)
    TC.ThePackageName = sPackageName
End Property

Public Property Get fileName() As String
    fileName = TC.fileName
End Property
Public Property Let fileName(sFileName As String)
    TC.fileName = sFileName
End Property

Public Property Get Path() As String
    Path = m_Path
End Property
Public Property Let Path(sPath As String)
    m_Path = sPath
End Property

Public Property Get CountSeries() As Integer
    CountSeries = m_Data.FieldCount
End Property

Public Property Get FieldName(iFld As Long) As String
    FieldName = m_Data.FieldName(iFld)
End Property
Public Property Let FieldName(iFld As Long, sFieldName As String)
    m_Data.FieldName(iFld) = sFieldName
End Property

Public Property Get Item(Index As Long, Field As Variant) As String
    Item = m_Data.Item(Index, Field)
End Property

Public Property Let Item(Index As Long, Field As Variant, sVal As String)
    m_Data.Item(Index, Field) = sVal
End Property
'======================================================================================================================================================
' Set Options from string
' The options example:
'     sOpt = "title=My Title^height=500^width=700^legend=[position=bottom;color=blue;fontsize=16]^chart=[title=THETITLE;subtutle=CRUTE]"
'======================================================================================================================================================
Public Function SetOptions(sOpt As String, Optional DLM As String = ";", Optional SEQ As String = "=", Optional SEP As String = vbCrLf) As Boolean
Dim sKey As String, sVal As String, sExt As String
Dim Arr() As String, I As Integer, nDim As Integer, iL As Integer

    On Error GoTo ErrHandle
'--------------------
    If sOpt = "" Then Exit Function
    If Left(sOpt, 1) = "{" Then            ' No need do something (it is already in Googke format)
        m_OPTIONS.Push sOpt
    Else
        Arr = Split(sOpt, SEP): nDim = UBound(Arr)
        For I = 0 To nDim
            iL = InStr(1, Arr(I), SEQ)
            If iL <> 0 Then
                    sKey = Trim(Left(Arr(I), iL - 1)): If sKey = "" Then GoTo nextOpt
                    sVal = Trim(Right(Arr(I), Len(Arr(I)) - iL))
             
                    If Left(sVal, 1) = "[" And Right(sVal, 1) = "]" Then
                         sVal = Trim(Mid(sVal, 2, Len(sVal) - 2))
                         AddOption sKey, , sVal
                    Else
                         AddOption sKey, sVal
                    End If
            End If
nextOpt:
        Next I
   End If
'--------------------
ExitHere:
    SetOptions = True '!!!!!!!!!!
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "SetOptions", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' Add Option/ For complex parameters, they can be specified directly in the options, which makes it possible to avoid parsing.
' The sign is the presence of {
'======================================================================================================================================================
Public Function AddOption(sKey As String, Optional sValue As String, Optional sParameters As String, Optional DLM As String = "=", _
                                                                                                               Optional SEP As String = ";") As Boolean
Dim sRes As String

    On Error GoTo ErrHandle
'--------------------
    If sKey = "" Then Exit Function
    
    If InStr(1, sValue, "{") > 0 Then                        ' MULTIPLE PARAMETERS ARE PLACED TO VALUE (nothing to do)
       sRes = sValue
    Else                                                     ' ONLY SINGLE VALUE AND PARAMETERS COULD BE STORE SEPARATELY
        If sParameters <> "" Then                            ' MULTIPLE PARAMETERS
            sRes = IIf(sValue <> "", sValue & ": ", "") & CreateParamForOption(sParameters, DLM, SEP)
        Else                                                 ' SINGLE OPTIONS
            sRes = IIf(IsNumeric(sValue), sValue, sCH(sValue))
        End If
    End If
    
    sRes = sKey & ": " & sRes
    
    m_OPTIONS.Push sRes
'--------------------
ExitHere:
    AddOption = True '!!!!!!!!!!!!
    Exit Function
'---------
ErrHandle:
    ErrPrint2 "AddOption", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Process parameters for options
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateParamForOption(sParameters As String, Optional DLM As String = "=", Optional SEP As String = ";") As String
Dim Arr() As String, nDim As Integer, sPairs() As String, sKey As String, sVal As String
Dim sRes As String, I As Integer

Const DPLTR As String = ": "
Const SPLTR As String = ", "

    On Error Resume Next
'------------------------------
Arr = Split(sParameters, SEP): nDim = UBound(Arr)
For I = 0 To nDim
      If InStr(1, Arr(I), DLM) > 0 Then
               sPairs = Split(Arr(I), DLM)
               sKey = sPairs(0): sVal = sPairs(1)
               If Not IsNumeric(sVal) Then sVal = sCH(sVal)
               Arr(I) = sKey & DPLTR & sVal     ' Reformat params
      End If
Next I
sRes = StrTag(Join(Arr, SPLTR & vbCrLf), "{ ", " }")
'------------------------------
ExitHere:
    CreateParamForOption = sRes '!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' The demo for Chart DONUTS
'======================================================================================================================================================
Public Function Demo_CHART_DONUTS() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_DONUTS
    If Me.fileName <> "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "DONULTS CHART" ': Me.ChartID = "donutchart"
    
    
    'Me.AddOption "title", Me.ChartTitle
    'Me.AddOption "pieHole", "0.4"
    
    Me.SetOptions "title=" & Me.ChartTitle & vbCrLf & "pieHole=0.4"
    
    Me.Width = 900: Me.height = 500

    ' FIELDS
    
    'Me.FieldName(0) = "Task": Me.FieldName(1) = "Hours per Day"
    Me.AddFieldList "Task;Hours per Day"
    
    ' VALUES
    Me.AddColumnData "Task", Array("Work", "Eat", "Commute", "Watch TV", "Sleep")
    Me.AddColumnData "Hours per Day", Array(11, 2, 2, 2, 7)
    
    'Me.Item(0, 0) = "Work": Me.Item(0, 1) = 11
    'Me.Item(1, 0) = "Eat": Me.Item(1, 1) = 2
    'Me.Item(2, 0) = "Commute": Me.Item(2, 1) = 2
    'Me.Item(3, 0) = "Watch TV": Me.Item(3, 1) = 2
    'Me.Item(4, 0) = "Sleep": Me.Item(4, 1) = 7
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_DONUTS = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_DONUTS", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'======================================================================================================================================================
' The demo for Chart BAR -2
'======================================================================================================================================================
Public Function Demo_CHART_BAR_2() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_BAR
    
    If Me.fileName <> "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "The Books": Me.ChartID = "barchart"
    
    'Me.AddOption "title", Me.ChartTitle
    'Me.AddOption "width", "600"
    'Me.AddOption "height", "400"
    
    'Me.AddOption "bar", , "groupWidth=95%"
    'Me.AddOption "legend", , "position=top;maxLines=3"
    'Me.AddOption "isStacked", "true"
    
    Me.Width = 900: Me.height = 300
    
    Me.SetOptions "title=" & Me.ChartTitle & vbCrLf & "width=" & 600 & vbCrLf & "height=" & 400 & vbCrLf & _
                  "bar=[groupWidth=95%]" & vbCrLf & "legend=[position=top;maxLines=3]" & vbCrLf & "isStacked=true"

    

    ' FIELDS
    
    'Me.FieldName(0) = "Genre": Me.FieldName(1) = "Fantasy & Sci Fi": Me.FieldName(2) = "Romance"
    'Me.FieldName(3) = "Mystery/Crime": Me.FieldName(4) = "General"
    'Me.FieldName(5) = "Western": Me.FieldName(6) = "Literature":  Me.FieldName(7) = "{ role: " & sCH("annotation") & " }"
    
    Me.AddFieldList "Genre;Fantasy & Sci Fi;Romance;Mystery/Crime;General;Western;Literature" & ";" & "ROLE_ANNOTATION"
    
    ' VALUES
    'Me.Item(0, 0) = sCH("2010"): Me.Item(0, 1) = 10: Me.Item(0, 2) = 24: Me.Item(0, 3) = 20: Me.Item(0, 4) = 32: Me.Item(0, 5) = 18: Me.Item(0, 6) = 5: Me.Item(0, 7) = sCH(" ")
    'Me.Item(1, 0) = sCH("2020"): Me.Item(1, 1) = 16: Me.Item(1, 2) = 22: Me.Item(1, 3) = 23: Me.Item(1, 4) = 30: Me.Item(1, 5) = 16: Me.Item(1, 6) = 9: Me.Item(1, 7) = sCH(" ")
    'Me.Item(2, 0) = sCH("2030"): Me.Item(2, 1) = 28: Me.Item(2, 2) = 19: Me.Item(2, 3) = 29: Me.Item(2, 4) = 30: Me.Item(2, 5) = 12: Me.Item(2, 6) = 13: Me.Item(2, 7) = sCH(" ")
    
    Me.AddColumnData "Genre", Array(sCH("2010"), sCH("2020"), sCH("2030"))
    Me.AddColumnData "Fantasy & Sci Fi", Array(10, 16, 20)
    Me.AddColumnData "Romance", Array(24, 22, 19)
    Me.AddColumnData "Mystery/Crime", Array(20, 23, 29)
    Me.AddColumnData "General", Array(32, 30, 30)
    Me.AddColumnData "Western", Array(18, 16, 12)
    Me.AddColumnData "Literature", Array(5, 9, 13)
    Me.AddColumnData "ROLE_ANNOTATION", Array("10", "20", "30")
    
    
    
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_BAR_2 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_BAR_2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The demo for Chart GAUGE
'======================================================================================================================================================
Public Function Demo_CHART_GAUGE() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_GAUGE
    

    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Computational Resources"  ': Me.ChartID = "columnchart_material"
    Me.Width = 400: Me.height = 120
    
    ' OPTIONS
    Me.SetOptions "chart=[title=" & Me.ChartTitle & vbCrLf & "width=400" & vbCrLf & "height=120" & vbCrLf & "redFrom=90" & vbCrLf & "yellowFrom=75" & vbCrLf & "yellowTo=90" & vbCrLf & "minorTick=5"
    
    ' FIELDS
    
    Me.AddFieldList "Label;Value"
    
    ' VALUES
    
    Me.AddColumnData "Label", Array("Memory", "CPU", "Network")
    Me.AddColumnData "Value", Array(80, 50, 68)
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_GAUGE = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_GAUGE", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The demo for Chart GAUGE
'======================================================================================================================================================
Public Function Demo_CHART_GEO() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_GEO
    

    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Colored Map"  ': Me.ChartID = "columnchart_material"
    Me.Width = 900: Me.height = 500
    
    ' OPTIONS
    Me.SetOptions "title=" & Me.ChartTitle
    ' FIELDS
    
    Me.AddFieldList "Country;Popularity"
    
    ' VALUES
    
    Me.AddColumnData "Country", Array("Germany", "United States", "Brazil", "Canada", "France", "RU")
    Me.AddColumnData "Popularity", Array(200, 300, 400, 500, 600, 700)
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_GEO = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_GEO", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'======================================================================================================================================================
' The demo for Chart COLUMNS
'======================================================================================================================================================
Public Function Demo_CHART_COLUMNS_3() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.IsMaterial = True       ' BEFORE SETTING CHARTNAME
    Me.ChartType = CHR_COLUMNS

    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Company Performance"  ': Me.ChartID = "columnchart_material"
    Me.Width = 800: Me.height = 500
    
    'Me.AddOption "chart", , "title=" & Me.ChartTitle & ";subtitle=Sales, Expenses, and Profit: 2014-2017"
    Me.SetOptions "chart=[title=" & Me.ChartTitle & ";subtitle=Sales, Expenses, and Profit: 2014-2017" & "]"
    
    ' FIELDS
    
    'Me.FieldName(0) = "Year": Me.FieldName(1) = "Sales": Me.FieldName(2) = "Expenses": Me.FieldName(3) = "Profit"
    Me.AddFieldList "Year;Sales;Expenses;Profit"
    
    ' VALUES
    'Me.Item(0, 0) = sCH("2014"): Me.Item(0, 1) = 1000: Me.Item(0, 2) = 400: Me.Item(0, 3) = 2000
    'Me.Item(1, 0) = sCH("2015"): Me.Item(1, 1) = 1170: Me.Item(1, 2) = 460: Me.Item(1, 3) = 250
    'Me.Item(2, 0) = sCH("2016"): Me.Item(2, 1) = 660: Me.Item(2, 2) = 1120: Me.Item(2, 3) = 300
    'Me.Item(3, 0) = sCH("2017"): Me.Item(3, 1) = 1030: Me.Item(3, 2) = 540: Me.Item(3, 3) = 350
    
    Me.AddColumnData "Year", Array(sCH("2014"), sCH("2015"), sCH("2016"), sCH("2017"))
    Me.AddColumnData "Sales", Array(1000, 1170, 660, 1030)
    Me.AddColumnData "Expenses", Array(400, 460, 1120, 540)
    Me.AddColumnData "Profit", Array(200, 250, 300, 350)
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_COLUMNS_3 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_COLUMNS_3", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'======================================================================================================================================================
' The demo for Chart COLUMNS
'======================================================================================================================================================
Public Function Demo_CHART_COLUMNS_2() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_COLUMNS

    
    If Me.fileName <> "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Book sales" ': Me.ChartID = "columnhart"
    
    Me.AddOption "title", Me.ChartTitle
    Me.AddOption "width", "600"
    Me.AddOption "height", "400"
    
 
    
    Me.SetOptions "title=" & Me.ChartTitle & vbCrLf & "width=600" & vbCrLf & "height=400" & vbCrLf & _
                  "bar=[groupWidth=75%]" & vbCrLf & "legend =[position=top;maxLines=3]" & vbCrLf & "isStacked=true"
    
    ' FIELDS
    
    Me.AddFieldList "Genre;Fantasy & Sci Fi;Romance;Mystery/Crime;General;Western;Literature;ROLE_ANNOTATION"
    
 
    Me.AddColumnData "Genre", Array(sCH("2010"), sCH("2020"), sCH("2030"))
    Me.AddColumnData "Fantasy & Sci Fi", Array(10, 16, 20)
    Me.AddColumnData "Romance", Array(24, 22, 19)
    Me.AddColumnData "Mystery/Crime", Array(20, 23, 29)
    Me.AddColumnData "General", Array(32, 30, 30)
    Me.AddColumnData "Western", Array(18, 16, 12)
    Me.AddColumnData "Literature", Array(5, 9, 13)
    Me.AddColumnData "ROLE_ANNOTATION", Array(" ", " ", " ")

    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_COLUMNS_2 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_COLUMNS_2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The demo for Chart COLUMNS
'======================================================================================================================================================
Public Function Demo_CHART_COLUMNS_1() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_COLUMNS

    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Density of Precious Metals, in g/cm^3" ': Me.ChartID = "columnhart"
    
    'Me.AddOption "title", Me.ChartTitle
    'Me.AddOption "width", "600"
    'Me.AddOption "height", "400"
    
    'Me.AddOption "bar", , "groupWidth=95%"
    'Me.AddOption "legend", "position=none"
    
    Me.SetOptions "title=" & Me.ChartTitle & vbCrLf & "width=600" & vbCrLf & "height=400" & vbCrLf & _
         "bar=[groupWidth=95%]" & vbCrLf & "legend=[position=none]"
    
    ' FIELDS
    
    'Me.FieldName(0) = "Element": Me.FieldName(1) = "Density": Me.FieldName(2) = "{ role: " & sCH("style") & " }": Me.FieldName(3) = "{ role: " & sCH("annotation") & " }"
    
    Me.AddFieldList "Element;Density;ROLE_STYLE;ROLE_ANNOTATION"

    ' VALUES
    'Me.Item(0, 0) = "Copper": Me.Item(0, 1) = 8.94: Me.Item(0, 2) = "#b87333": Me.Item(0, 3) = "Cu"
    'Me.Item(1, 0) = "Silver": Me.Item(1, 1) = 10.49: Me.Item(1, 2) = "silver": Me.Item(1, 3) = "Ag"
    'Me.Item(2, 0) = "Gold": Me.Item(2, 1) = 19.3: Me.Item(2, 2) = "gold": Me.Item(2, 3) = "Au"
    'Me.Item(3, 0) = "Platinum": Me.Item(3, 1) = 21.45: Me.Item(3, 2) = "color: #e5e4e2": Me.Item(3, 3) = "Pt"
    
    Me.AddColumnData "Element", Array("Copper", "Silver", "Gold", "Platinum")
    Me.AddColumnData "Density", Array(8.94, 10.49, 19.3, 21.45)
    Me.AddColumnData "ROLE_STYLE", Array("#b87333", "silver", "gold", "color: #e5e4e2")
    Me.AddColumnData "ROLE_ANNOTATION", Array("Cu", "Ag", "Au", "Pt")
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_COLUMNS_1 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_COLUMNS_1", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The demo for Gantt Chart
'======================================================================================================================================================
Public Function Demo_CHART_GANTT_2() As String
Dim sRes As String
Dim TASKS() As TTASK, nDim As Integer, I As Integer
Dim TaskIDs As Variant, TaskNames As Variant, RESOURCES As Variant
Dim StartDates As Variant, EndDates As Variant, Durations As Variant, Dependencies As Variant, Completeness As Variant

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_GANTT
    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Fertility rate vs life expectancy in selected countries (2010)"
    
   ' OPTIONS
    Me.SetOptions "height=275"
    
    Me.Width = 900: Me.height = 500
   
   
   ' CREATE TASK STACK
    TaskIDs = Array("Research", "Write", "Cite", "Complete", "Outline")
    TaskNames = Array("Find sources", "Write paper", "Create bibliography", "Hand in paper", "Outline pape")
    RESOURCES = Array("", "write", "write", "complete", "write")
    StartDates = Array("2015-01-01", "", "", "", "")
    EndDates = Array("2015-01-05", "2015-01-09", "2015-01-07", "2015-01-10", "2015-01-06")
    Durations = Array("", daysToMilliseconds(3), daysToMilliseconds(1), daysToMilliseconds(1), daysToMilliseconds(1))
    Completeness = Array(100, 25, 20, 0, 100)
    Dependencies = Array("", "Research,Outline", "Research", "Cite,Write", "Research")
    
    nDim = UBound(TaskIDs): ReDim TASKS(nDim)
    For I = 0 To nDim
        
        TASKS(I).TaskID = TaskIDs(I): TASKS(I).TaskName = TaskNames(I): TASKS(I).Resource = FormatNull(RESOURCES(I))
        
        TASKS(I).StartDate = FormatNull(StartDates(I)): TASKS(I).EndDate = FormatNull(EndDates(I))
        TASKS(I).Duration = FormatNull(Durations(I)): TASKS(I).PercentComplete = FormatNull(Completeness(I))
        TASKS(I).Dependencies = FormatNull(Dependencies(I))
    Next I
    
    m_sEXT = FormatTaskColumns & vbCrLf & GetTaskDataStr(TASKS)  ' STORE ADDITIONAL INFO FOR NEXT USAGE. LATER COULD REMOVE PRIVATE TYPE TTASK
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_GANTT_2 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_GANTT_2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Coun of millisecund for example
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function daysToMilliseconds(nDays As Long) As Long
      daysToMilliseconds = nDays * 24 * 60 * 60 * 1000
End Function
'======================================================================================================================================================
' The demo for Gantt Chart
'======================================================================================================================================================
Public Function Demo_CHART_GANTT_1() As String
Dim sRes As String, sOpt As String
Dim TASKS() As TTASK, nDim As Integer, I As Integer
Dim TaskIDs As Variant, TaskNames As Variant, RESOURCES As Variant
Dim StartDates As Variant, EndDates As Variant, Durations As Variant, Dependencies As Variant, Completeness As Variant

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_GANTT
    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Fertility rate vs life expectancy in selected countries (2010)"
    
   ' OPTIONS
    Me.SetOptions "height=400" & vbCrLf & "gantt=[trackHeight=30]"
    
    Me.Width = 900: Me.height = 500
   
   
   ' CREATE TASK STACK
    TaskIDs = Array("2014Spring", "2014Summer", "2014Autumn", "2014Winter", "2015Spring", "2015Summer", "2015Autumn", "2015Winter", "Football", "Baseball", "Basketball", "Hockey")
    TaskNames = Array("Spring 2014", "Summer 2014", "Autumn 2014", "Winter 2014", "Spring 2015", "Summer 2015", "Autumn 2015", "Winter 2015", "Football Season", "Baseball Season", "Basketball Season", "Hockey Season")
    RESOURCES = Array("spring", "summer", "autumn", "winter", "spring", "summer", "autumn", "winter", "sports", "sports", "sports", "sports")
    StartDates = Array("2014-02-22", "2014-05-21", "2014-08-21", "2014-11-21", "2015-02-22", "2015-05-21", "2015-08-21", "2015-11-21", "2014-08-04", "2015-02-28", "2014-09-28", "2014-09-08")
    EndDates = Array("2014-05-20", "2014-08-20", "2014-11-20", "2015-02-21", "2015-05-20", "2015-08-20", "2015-11-20", "2016-02-21", "2015-01-01", "2015-09-20", "2015-05-20", "2015-05-21")
    Durations = Array("", "", "", "", "", "", "", "", "", "", "", "")
    Completeness = Array(100, 100, 100, 100, 50, 0, 0, 0, 100, 14, 86, 89)
    Dependencies = Array("", "", "", "", "", "", "", "", "", "", "", "")
    
    nDim = UBound(TaskIDs): ReDim TASKS(nDim)
    For I = 0 To nDim
        
        TASKS(I).TaskID = TaskIDs(I): TASKS(I).TaskName = TaskNames(I): TASKS(I).Resource = FormatNull(RESOURCES(I))
        
        TASKS(I).StartDate = FormatNull(StartDates(I)): TASKS(I).EndDate = FormatNull(EndDates(I))
        TASKS(I).Duration = FormatNull(Durations(I)): TASKS(I).PercentComplete = FormatNull(Completeness(I))
        TASKS(I).Dependencies = FormatNull(Dependencies(I))
    Next I
    
    m_sEXT = FormatTaskColumns & vbCrLf & GetTaskDataStr(TASKS)  ' STORE ADDITIONAL INFO FOR NEXT USAGE. LATER COULD REMOVE PRIVATE TYPE TTASK
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_GANTT_1 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_GANTT_1", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Format "-" as null
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatNull(str As Variant) As String
Dim sWork As String
    sWork = CStr(str)
    FormatNull = IIf(sWork <> "", sWork, "null")   '!!!!!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Columns for Gantt Chart
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatTaskColumns() As String
Dim sRes As String

Const nShoft As Integer = 7

    On Error Resume Next
'-------------------
      sRes = AppendString("", "var data = new google.visualization.DataTable();", vbCrLf, nShoft)

      sRes = AppendString(sRes, "data.addColumn(" & sCH("string") & ", " & sCH("Task ID") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("string") & ", " & sCH("Task Name") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("string") & ", " & sCH("Resource") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("date") & ", " & sCH("Start Date") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("date") & ", " & sCH("End Date") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("number") & ", " & sCH("Duration") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("number") & ", " & sCH("Percent Completen") & ");", vbCrLf, nShoft + 3)
      sRes = AppendString(sRes, "data.addColumn(" & sCH("string") & ", " & sCH("Dependencies") & ");", vbCrLf, nShoft + 3)

'-------------------
ExitHere:
     FormatTaskColumns = sRes '!!!!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Data Tasks string presentation
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetTaskDataStr(TASKS() As TTASK) As String
Dim nDim As Integer, sRes As String, I As Integer
Dim sWork As String

Const nShoft As Long = 7
Const SPLT As String = ", "

    On Error Resume Next
'------------------------------
sRes = AppendString("", "data.addRows([", vbCrLf, nShoft)

nDim = UBound(TASKS)
For I = 0 To nDim
    'Debug.Print TASKS(I).TaskID
    sWork = IsSkipSch(TASKS(I).TaskID) & SPLT & IsSkipSch(TASKS(I).TaskName) & SPLT & IsSkipSch(TASKS(I).Resource) & SPLT
    sWork = sWork & FormatDate2(TASKS(I).StartDate) & SPLT & FormatDate2(TASKS(I).EndDate) & SPLT
    sWork = sWork & IsSkipSch(TASKS(I).Duration) & SPLT & IsSkipSch(TASKS(I).PercentComplete) & SPLT & IsSkipSch(TASKS(I).Dependencies)

    
    sRes = AppendString(sRes, SHT(sWork) & SPLT, vbCrLf, nShoft + 3)
Next I
sRes = AppendString(sRes, "]);", vbCrLf, nShoft)
'------------------------------
ExitHere:
    GetTaskDataStr = sRes '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Check - if Null or Number - scip sCh
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function IsSkipSch(str As String) As String
Dim sRes As String
    
    On Error GoTo ErrHandle
'----------------
    If Left(str, 1) = Chr(39) Then
        sRes = str
    ElseIf IsNumeric(str) Or str = "null" Then
        sRes = str
    Else
        sRes = sCH(str)
    End If
'----------------
ExitHere:
    IsSkipSch = sRes '!!!!!!!!!!!!
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "IsSkipSch", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Format Date
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatDate2(sDate As String) As String
Dim sRes As String

Const SPLT As String = ", "

    On Error GoTo ErrHandle
'----------------------
If sDate = "null" Or sDate = "" Then
    sRes = "null"
ElseIf IsDate(sDate) Then
    sRes = "new Date(" & Year(sDate) & SPLT & Month(sDate) & SPLT & Day(sDate) & ")"
Else
    Err.Raise 10007, , "Can' recognize date"
End If
'----------------------
ExitHere:
    FormatDate2 = sRes  '!!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "FormatDate2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' The demo for Bubble Chart
'======================================================================================================================================================
Public Function Demo_CHART_BUBBLE_2() As String
Dim sRes As String, sOpt As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_BUBBLE
    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Fertility rate vs life expectancy in selected countries (2010)"
    
   ' OPTIONS
    Me.SetOptions "colorAxis={colors: [" & sCH("yellow") & "," & sCH("red") & "]}"
    
    Me.Width = 900: Me.height = 500

    ' FIELDS
    
    Me.AddFieldList "ID;X;Y;Temperature"

    ' VALUES
        
    Me.AddColumnData "ID", Array(" ", " ", " ", " ", " ", " ", " ")
    Me.AddColumnData "X", Array(80, 79, 78, 72, 81, 72, 68)
    Me.AddColumnData "Y", Array(167, 136, 184, 278, 200, 170, 477)
    Me.AddColumnData "Temperature", Array(120, 130, 50, 230, 210, 100, 80)
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_BUBBLE_2 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_BUBBLE_2", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The demo for Bubble Chart
'======================================================================================================================================================
Public Function Demo_CHART_BUBBLE_1() As String
Dim sRes As String, sOpt As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_BUBBLE
    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Fertility rate vs life expectancy in selected countries (2010)"
    
    Me.AddStrToOPt "title=" & Me.ChartTitle & ", " & " X=Life Expectancy, Y=Fertility, Bubble size=Population, Bubble color=Region"
    Me.AddStrToOPt "hAxis=[title=Life Expectancy]"
    Me.AddStrToOPt "vAxis=[title=Fertility Rate]"
    sOpt = Me.AddStrToOPt("bubble=[textStyle={fontSize:11}]")
    
    Me.SetOptions sOpt
    
    Me.Width = 900: Me.height = 500

    ' FIELDS
    
    Me.AddFieldList "ID;Life Expectancy;Fertility Rate;Region;Population"

    ' VALUES
        
    Me.AddColumnData "ID", Array("CAN", "DEU", "DNK", "EGY", "GBR", "IRN", "IRQ", "ISR", "RUS", "USA")
    Me.AddColumnData "Life Expectancy", Array(80.66, 79.84, 78.6, 72.73, 80.05, 72.49, 68.09, 81.55, 68.6, 78.09)
    Me.AddColumnData "Fertility Rate", Array(1.67, 1.36, 1.84, 2.78, 2, 1.7, 4.77, 2.96, 1.54, 2.05)
    Me.AddColumnData "Region", Array("North America", "Europe", "Europe", "Middle East", "Europe", "Middle East", "Middle East", "Middle East", "Europe", "North America")
    Me.AddColumnData "Population", Array(33739900, 81902307, 5523095, 79716203, 61801570, 73137148, 31090763, 7485600, 141850000, 307007000)
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_BUBBLE_1 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_BUBBLE_1", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'======================================================================================================================================================
' The Options
'======================================================================================================================================================
Public Function AddStrToOPt(Optional sNewOption As String) As String
Dim sRes As String
    m_sOPT = AppendString(m_sOPT, sNewOption, vbCrLf, 0)
'-----------------
ExitHere:
    AddStrToOPt = m_sOPT '!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' Create Bar Chart: The main intrance to create chart bar:                                                                                      |||||||
' sFile - the location to save the html file after the diagram is generated, sTitle - chart title                                            ||||||||||
' sFields - list of columns, the first one is category (x-axis), sOptions - the options for Google Charts,                                          |||
' X = - 1-d array with category var, Y - 2D array of other columns (include ROLE_STYLE and ROLE_ANNOTATION), width, height - size of frame|||||||||||||
'=======================================================================================================================================================
Public Function CreateChartBar(Optional sPath As String, Optional sTitle As String = "Bar Chart", Optional sOptions As String, _
                               Optional sFields As String = "X;Y1;Y2", Optional x As Variant, Optional y As Variant, _
                                                 Optional Width As Long = 600, Optional height As Long = 400, Optional DLM As String = ";") As String
Dim sRes As String, FLDS() As String, nFlds As Integer, J As Integer
Dim I As Integer, nColumns As Integer, nRows As Integer

    On Error Resume Next
'------------------------
    If sPath <> "" Then
        Me.Folder = FolderNameOnlyAbstract(sPath): Me.fileName = FileNameOnly(sPath)
    Else
        Me.fileName = CreateOutputFileName()
        Me.Folder = CreateOutPutFolder()
    End If
    
    Me.Width = Width: Me.height = height             ' SET FRAME SIZE
    Me.SetOptions sOptions, DLM, "=", vbCrLf         ' SET THE OPTIONS
    Me.AddFieldList sFields                          ' SET FIELD LIST
    
    FLDS = Split(sFields, DLM): nFlds = UBound(FLDS)
    nColumns = UBound(y, 2): If nFlds <> nColumns + 1 Then Err.Raise 10008, , "Wrong field number"
    
    
'------------------------
ExitHere:
    CreateChartBar = sRes '!!!!!!!!!!!!!!!!!
End Function
'======================================================================================================================================================
' The demo for Chart BAR
'======================================================================================================================================================
Public Function Demo_CHART_BAR_1() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_BAR
    
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    Me.ChartTitle = "Density of Precious Metals, in g/cm^3"  ': Me.ChartID = "barchart"
    
    'Me.AddOption "title", Me.ChartTitle
    'Me.AddOption "width", "600"
    'Me.AddOption "height", "400"
    
    
   'Me.AddOption "bar", , "groupWidth=95%"
   'Me.AddOption "legend", , "position=none"
    
    Me.SetOptions "title=" & Me.ChartTitle & vbCrLf & "width=900" & vbCrLf & "height=400" & vbCrLf & "bar=" & "[groupWidth=95%]" & vbCrLf & "legend=[position=none]"
    
    
    Me.Width = 900: Me.height = 300

    ' FIELDS
    
    'Me.FieldName(0) = "Element": Me.FieldName(1) = "Density": Me.FieldName(2) = "{ role: " & sCH("style") & " }": Me.FieldName(3) = "{ role: " & sCH("annotation") & " }"
    Me.AddFieldList "Element;Density;" & "ROLE_STYLE" & ";" & "ROLE_ANNOTATION"

    ' VALUES

    
    Me.AddColumnData "Element", Array("Copper", "Silver", "Gold", "Platinum")
    Me.AddColumnData "Density", Array(8.94, 10.49, 19.3, 21.45)
    Me.AddColumnData "ROLE_STYLE", Array("#b87333", "silver", "gold", "color: #e5e4e2")
    Me.AddColumnData "ROLE_ANNOTATION", Array("Cu", "Ag", "Au", "Pt")
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_BAR_1 = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_BAR_1", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Add Fields by String
'======================================================================================================================================================
Public Function AddFieldList(sFields As String, Optional DLM As String = ";") As Boolean
Dim Arr() As String, nDim As Long, I As Long

    On Error GoTo ErrHandle
'--------------------------
    If sFields = "" Then Exit Function
    Arr = Split(sFields, DLM): nDim = UBound(Arr)
    For I = 0 To nDim
        m_Data.FieldName(I) = Arr(I)
    Next I
'--------------------------
ExitHere:
    AddFieldList = True '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "AddFieldList", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' Add Column Data: vData Array or Delimited String
'======================================================================================================================================================
Public Function AddColumnData(sField As String, vData As Variant, Optional DLM As String) As Boolean
Dim I As Long, nDim As Long, sARR() As String, sVal As String
    On Error GoTo ErrHandle
'----------------------------
    If IsArray(vData) Then
        nDim = UBound(vData)
        For I = 0 To nDim
              sVal = CStr(vData(I))
              m_Data.Item(I, sField) = sVal
        Next I
    Else
        sARR = Split(CStr(vData), DLM): nDim = UBound(sARR)
        For I = 0 To nDim
              sVal = sARR(I)
              m_Data.Item(I, sField) = sVal
        Next I
    End If
'----------------------------
ExitHere:
    AddColumnData = True '!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "AddColumnData", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' The demo for Chart Line
'======================================================================================================================================================
Public Function Demo_CHART_LINE() As String
Dim sRes As String

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_LINE
    Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    Me.ChartTitle = "Some functions": Me.ChartID = "curve_chart"
    
    'Me.AddOption "title", Me.ChartTitle
    'Me.AddOption "curveType", "function"
    'Me.AddOption "legend", , "position=bottom"
    Me.SetOptions ("title=" & Me.ChartTitle & vbCrLf & "curveType=function" & vbCrLf & "legend=[position=bottom]")
    
    Me.Width = 500: Me.height = 600

    ' FIELDS
    Me.AddFieldList "X;X^2;X^3"
    
    'Me.FieldName(0) = "X": Me.FieldName(1) = "X^2": Me.FieldName(2) = "X^3"

    ' VALUES
    Me.AddColumnData "X", Array(-3, -2, -1, 0, 1, 2, 3)
    Me.AddColumnData "X^2", Array(9, 4, 1, 0, 1, 4, 9)
    Me.AddColumnData "X^3", Array(-27, -8, -1, 0, 1, 8, 27)
      
    'Me.Item(0, 0) = -3: Me.Item(0, 1) = 9: Me.Item(0, 2) = -27
    'Me.Item(1, 0) = -2: Me.Item(1, 1) = 4: Me.Item(1, 2) = -8
    'Me.Item(2, 0) = -1: Me.Item(2, 1) = 1: Me.Item(2, 2) = -1
    'Me.Item(3, 0) = 0: Me.Item(3, 1) = 0: Me.Item(3, 2) = 0
    'Me.Item(4, 0) = 1: Me.Item(4, 1) = 1: Me.Item(4, 2) = 1
    'Me.Item(5, 0) = 2: Me.Item(5, 1) = 4: Me.Item(5, 2) = 8
    'Me.Item(6, 0) = 3: Me.Item(6, 1) = 9: Me.Item(6, 2) = 27
    
    
    If Build(Me.ChartType) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_LINE = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_LINE", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'======================================================================================================================================================
' The demo for Public Chart for Pie type
'======================================================================================================================================================
Public Function Demo_CHART_PIE() As String
Dim sRes As String

Const DLM As String = ";"

    On Error GoTo ErrHandle
'-----------------------
    Me.ChartType = CHR_PIE
    If Me.fileName = "" Then Me.fileName = CreateOutputFileName()
    Me.Folder = CreateOutPutFolder()
    
    
    Me.ChartTitle = "Montly sale of Books": Me.ChartID = "b_sale"
    
    'OPTIONS
    Call SetOptions("title=" & Me.ChartTitle)
    
    Me.height = 500: Me.Width = 700
    
    
    ' FIELDS
    Me.AddFieldList "Month;Sales Figure", DLM
    

    
    
    ' DATA/CATEGORY (X-AXIS)
    Me.AddColumnData "Month", Array("JAN", "FEB", "MAR", "APR", "MAY", "JUN", "JUL", "AUG", "SEP", "OCT", "NOV", "DEC")
    
   
        
    ' SERIES (YAXIS
    Me.AddColumnData "Sales Figure", Array(21, 32, 19, 8, 26, 31, 37, 16, 25, 23, 28, 17)
    
    
   
    
    
    If Build(CHR_PIE) Then sRes = SaveChart()
'-----------------------
ExitHere:
    Demo_CHART_PIE = sRes '!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Demo_CHART_PIE", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function



'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Private Sub Class_Initialize()

Const DLM As String = ";"
      
      Set m_Data = New cVTable
      Set m_OPTIONS = New cStack
      
      m_Data.AddFieldList "X_AXIS;Y1_AXIS", DLM
      
      
      m_ChartType = CHR_PIE
      Call CreateConfig(m_ChartType)
      
      TC.ThePackageName = "corechart"
End Sub

Private Sub Class_Terminate()
      Set m_Data = Nothing
      Set m_OPTIONS = Nothing
End Sub

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Function Build The sring that represent chart
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function Build(ChartType As TCHARTTYPE) As Boolean
Dim sWork As String

Const DLM As String = vbCrLf
Const nShoft As Long = 4
Const nShoftPlus As Long = 7

    On Error GoTo ErrHandle
'---------------------
    
    Select Case ChartType
    Case CHR_PIE:  ' (SPECIAL CASE)
    
         '<HEAD> GENERATION
         sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
         sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
         m_Res = GetSimpleTag("head", sWork, DLM, 0) & vbCrLf
         
         
         ' <BODY> GENERATION
         sWork = AppendString("", "<div id=" & QR(TC.ChartID) & " style=" & QR("width:" & m_Width & _
                                                 "px; height:" & m_Height & "px;") & "></div>", DLM, nShoft)
         sWork = GetSimpleTag("body", sWork, DLM, 0)
         m_Res = AppendString(m_Res, sWork, DLM) & vbCrLf
         
         
         ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
         sWork = AppendString("", BuildVizAPI(TC.ChartName, TC.DrawFuncName), DLM, nShoft)
         sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
         sWork = GetSimpleTag("script", sWork, DLM, 0)
         m_Res = AppendString(m_Res, sWork, DLM)
    
    Case CHR_LINE:
         
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
           
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf  '!!!!!!!!!!!!
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          ' <BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_DONUTS:
    
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, GetLoader, DLM, nShoft)
          
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          '<BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_BAR:
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, GetLoader, DLM, nShoft)
          
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_COLUMNS:
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, GetLoader, DLM, nShoft)
          
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_GAUGE:
              
              '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
           
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf  '!!!!!!!!!!!!
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          ' <BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_BUBBLE:
          
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
           
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf  '!!!!!!!!!!!!
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          ' <BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
    
    Case CHR_GANTT:        ' (SPECIAL CASE)
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
           
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf  '!!!!!!!!!!!!
          
          sWork = AppendString(sWork, BuildFunctionGantt(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          ' <BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)
        
    Case CHR_GEO:
          ' THIS IS REALIZE ONLY CASE WITHOUT MAP_KEY
              
          '<HEAD> GENERATION
          sWork = String(nShoft, " ") & AppendString("", GetSimpleTag("title", "GRACKLE GENERATED: " & TC.ChartName, "", 0), nShoft)
          sWork = AppendString(sWork, TC.APIProvider, DLM, nShoft)
           
          ' <SCRIPT> GENERATION (INCLUDING FUNC BODY)
          sWork = AppendString(sWork, "<script type=" & QR("text/javascript") & ">", DLM, nShoft)
          sWork = AppendString(sWork, BuildVizAPI(TC.ChartName, TC.DrawFuncName, 1), DLM, nShoft) & vbCrLf & vbCrLf  '!!!!!!!!!!!!
          sWork = AppendString(sWork, BuildFunction(TC.DrawFuncName, TC.ChartID, "data"), DLM, nShoft)
          sWork = AppendString(sWork, "</script>", DLM, nShoft)
          
          m_Res = GetSimpleTag("head", sWork, DLM, nShoft)  ' Complete <head>
          
          ' <BODY> GENERATION
          m_Res = AppendString(m_Res, GetSimpleTag("body", "<div id=" & QR(TC.ChartID) & " style=" & QR("width: " & m_Width & "px; height: " & m_Height & "px") & "></div>"), DLM, nShoft)

    Case Else:
    
    End Select
    
m_Res = "<!doctype html>" & vbCrLf & GetSimpleTag("html", m_Res, DLM, 0)
'-----------------------------
ExitHere:
    Build = True '!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "Build", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build API rows
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildVizAPI(sChartName As String, sFuncName As String, Optional VizAPIVer As Integer = 0) As String
Dim sRes As String

Const nShoft As Long = 4
Const DLM As String = vbCrLf

      On Error Resume Next
'----------------------------
         sRes = AppendString("", "// VISUALIZATION API AND THE " & sChartName & " PACKAGE.", DLM, nShoft)
      Select Case VizAPIVer:
      Case 0:
         sRes = AppendString(sRes, "google.load(" & QR("visualization") & ", " & QR("1") & ", { packages: " & SHT(QR("corechart")) & " });", DLM, nShoft)
         sRes = AppendString(sRes, "google.setOnLoadCallback(" & sFuncName & ");", vbCrLf, nShoft) & DLM
      Case 1:
         sRes = AppendString(sRes, "google.charts.load(" & sCH("current") & ", {" & sCH("packages") & ":[" & sCH(Me.Package) & "]});", DLM, nShoft)
         sRes = AppendString(sRes, "google.charts.setOnLoadCallback(" & sFuncName & ");", DLM, nShoft)
      Case Else
      End Select
'----------------------------
ExitHere:
         BuildVizAPI = sRes '!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildFunction(sFuncName As String, ChartID As String, Optional sdataview As String = "data") As String
Dim sRes As String, iVersion As Integer

Const nShoft As Long = 7
Const nShoftMinus As Long = 4

    On Error Resume Next
'---------------------
    sRes = "function " & sFuncName & "() {"
    sRes = AppendString(sRes, BuildOptions(), vbCrLf, nShoft) & vbCrLf
    
    Select Case sdataview
    Case "data":
        sRes = AppendString(sRes, BuildData(), vbCrLf, nShoft)
    Case "view":
    End Select
    
    If Me.IsMaterial Then iVersion = 1
    sRes = AppendString(sRes, BuildDraw(sdataview, sFuncName, ChartID, iVersion), vbCrLf, nShoft)
    sRes = AppendString(sRes, "}", vbCrLf, nShoftMinus)
'---------------------
ExitHere:
    BuildFunction = sRes '!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build Function For Gantt
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildFunctionGantt(sFuncName As String, ChartID As String, Optional sdataview As String = "data") As String
Dim sRes As String, iVersion As Integer, FLDS() As String, nFlds As Integer, I As Integer
Dim nRecords As Long, J As Long, sRecord As String, sData As String
Dim FldType As String

Const nShoft As Long = 7
Const nShoftMinus As Long = 4
Const SPLT As String = ", ", DLM As String = ";"

    On Error Resume Next
'---------------------
    sRes = "function " & sFuncName & "() {"
    
    ' OPTIONS IS THE SAME
    sRes = AppendString(sRes, BuildOptions(), vbCrLf, nShoft) & vbCrLf
    
    '------------------------------------------------------- CREATE DATA
    sRes = AppendString(sRes, m_sEXT, vbCrLf, 0)
        
    sRes = AppendString(sRes, BuildDraw(sdataview, sFuncName, ChartID, 0), vbCrLf, nShoft)
    sRes = AppendString(sRes, "}", vbCrLf, nShoftMinus)
'---------------------
ExitHere:
    BuildFunctionGantt = sRes '!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Google jsapi
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetJSAPI(Optional nShift As Long = 4) As String
Dim sShift As String
    If nShift > 0 Then sShift = String(nShift, " ")
    GetJSAPI = sShift & "<script type=" & QR("text/javascript") & " src=" & QR("https://www.google.com/jsapi") & "></script>"
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Google script loader
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetLoader(Optional nShift As Long = 4) As String
Dim sShift As String
    If nShift > 0 Then sShift = String(nShift, " ")
    GetLoader = sShift & "<script type=" & QR("text/javascript") & " src=" & QR("https://www.gstatic.com/charts/loader.js") & "></script>"
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Simple Tag
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetSimpleTag(sTag As String, sBodyTag As String, Optional DLM As String = vbCrLf, Optional nShift As Long = 4) As String
Dim sShift As String
    If nShift > 0 Then sShift = String(nShift, " ")
    
    GetSimpleTag = sShift & "<" & sTag & ">" & DLM & sBodyTag & DLM & sShift & "</" & sTag & ">"
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build Chart Draw
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildDraw(Optional sdataview As String = "data", Optional sFuncName As String = "PieChart", Optional ChartID As String, _
                                                                                                             Optional FuncVer As Integer = 0) As String
Dim sRes As String

Const nShoft As Long = 10

    On Error Resume Next
'---------------------------
Select Case FuncVer
Case 0:
    sRes = AppendString("", "var chart = new google.visualization." & TC.TheClass & "(document.getElementById(" & sCH(ChartID) & " ));", vbCrLf, nShoft)
    sRes = AppendString(sRes, "chart.draw(" & sdataview & ", options);", vbCrLf, nShoft)
Case 1:
    sRes = AppendString("", "var chart = new google.charts." & TC.TheClass & "(document.getElementById(" & sCH(ChartID) & "));", vbCrLf, nShoft)
    sRes = AppendString(sRes, "chart.draw(" & sdataview & ", google.charts." & TC.TheClass & ".convertOptions (options));", vbCrLf, nShoft)
End Select
'-----------------------
ExitHere:
    BuildDraw = sRes '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build Full Data String
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildData() As String
Dim sRes As String

Const nShoft As Long = 10

    On Error Resume Next
'---------------------------

sRes = AppendString("", "// CREATE A DataTable AND ADD DATA TO IT.", vbCrLf, nShoft)
sRes = AppendString(sRes, "var data = google.visualization.arrayToDataTable([", vbCrLf, nShoft)
sRes = AppendString(sRes, GetDataStr(), vbCrLf, nShoft + 3)
sRes = AppendString(sRes, "]);", vbCrLf, nShoft)
'-----------------------
ExitHere:
    BuildData = sRes '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Build Full Option String
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function BuildOptions() As String
Dim sRes As String

Const nShoft As Long = 10

    On Error Resume Next
'------------------------
    sRes = AppendString("", "var options = {", vbCrLf, nShoft)
    sRes = AppendString(sRes, GetOptionStr, vbCrLf)
    sRes = AppendString(sRes, "};", vbCrLf)
'-----------------------
ExitHere:
    BuildOptions = sRes '!!!!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Get Options String
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetOptionStr() As String
Dim sRes As String, sShift As String
Dim DLM As String

Const SPLT As String = ","
Const nShoft As Long = 12

    On Error Resume Next
'----------------
    If m_OPTIONS.Count = 0 Then Exit Function
    sShift = String(nShoft, " "): DLM = SPLT & vbCrLf & sShift
    sRes = sShift & m_OPTIONS.ToString(DLM)
'----------------
ExitHere:
    GetOptionStr = sRes '!!!!!!!!!!!
End Function
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Convert Data to string
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetDataStr() As String
Dim sRes As String, nFld As Integer, nRecord As Long
Dim sFlds As String, I As Long, sWork As String

Const SPLT As String = ",", DLM As String = ";"
Const nArrayShift As Integer = 10, nAllShift As Integer = 7
Const PreData As String = "// CREATE A DataTable AND ADD DATA TO IT."
    
    On Error GoTo ErrHandle
'--------------------
nFld = m_Data.FieldCount - 1: nRecord = m_Data.Count - 1
If nFld < 0 Or nRecord < 0 Then Exit Function

sFlds = m_Data.GetFldString(DLM): If sFlds = "" Then Exit Function
sFlds = FormatCells(sFlds, DLM, True)         ' And Filtre special Fields

sRes = AppendString(sRes, sFlds, nArrayShift)

For I = 0 To nRecord
    sWork = Join(m_Data.GetRow(I), DLM)
    If sWork <> "" Then
          sWork = FormatCells(sWork, DLM)
          sRes = AppendString(sRes, sWork, SPLT & vbCrLf, nArrayShift)
    End If
Next I

'---------------------
ExitHere:
    GetDataStr = sRes '!!!!!!!!
    Exit Function
'--------
ErrHandle:
    ErrPrint2 "GetDataStr", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Format Cell String/Number
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function FormatCells(sRow As String, Optional DLM As String = ";", Optional IsFld As Boolean) As String
Dim Arr() As String, nDim As Integer, I As Integer, sWork As String
Dim sRes As String

Const SPLT As String = ","

    On Error GoTo ErrHandle
'----------------------------
If sRow = "" Then Exit Function
Arr = Split(sRow, DLM): nDim = UBound(Arr)
For I = 0 To nDim
     sWork = Arr(I)
     If IsFld Then sWork = GetFieldReplacement(sWork)   ' For field check if special fld then replace it
     
     If sWork <> "" Then
        If IsNumeric(sWork) Then
           Arr(I) = Replace(Arr(I), ",", ".")
        ElseIf InStr(1, sWork, "{") Then
           Arr(I) = sWork
        ElseIf sWork = "null" Then
           Arr(I) = sWork
        Else
            If Left(Arr(I), 1) <> Chr(39) Then Arr(I) = sCH(Arr(I))
        End If
     End If
Next I
sRes = SHT(Join(Arr, SPLT))
'----------------------------
ExitHere:
    FormatCells = sRes '!!!!!!!!!!!!!!!
    Exit Function
'----------
ErrHandle:
    ErrPrint2 "FormatCells", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Output Folder
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateOutPutFolder() As String
Dim sRes As String

Const GEN_OUTPUT As String = "CHARTS"

    On Error GoTo ErrHandle
'---------------------
     sRes = CurrentProject.Path & "\" & GEN_OUTPUT
     If Dir(sRes, vbDirectory) = "" Then
         If Not FolderCreate(sRes) Then Err.Raise 10007, , "Can't create folder " & sRes
     End If
     
'---------------------
ExitHere:
    CreateOutPutFolder = sRes '!!!!!!!!!!!!
    Exit Function
'-----
ErrHandle:
    ErrPrint2 "CreateOutPutFolder", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function


'------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Output FileName
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function CreateOutputFileName() As String
Dim sFile As String

    On Error Resume Next
'-------------------
    Select Case m_ChartType
         Case CHR_BAR: sFile = "CHR_BAR"
         Case CHR_BUBBLE: sFile = "CHR_BUBBLE"
         Case CHR_COLUMNS: sFile = "CHR_COLUMNS"
         Case CHR_DONUTS: sFile = "CHR_DONUTS"
         Case CHR_GANTT: sFile = "CHR_GANTT"
         Case CHR_GAUGE: sFile = "CHR_GAUGE"
         Case CHR_GEO: sFile = "CHR_GEO"
         Case CHR_LINE: sFile = "CHR_LINE"
         Case CHR_PIE: sFile = "CHR_PIE"
         Case CHR_TIMELINES: sFile = "CHR_TIMELINES"
         Case Else: sFile = "CHR_OTHER"
    End Select
    
    sFile = Format(Now(), "yyyymmdd") & "_" & sFile & ".html"
'-------------------
ExitHere:
    CreateOutputFileName = sFile '!!!!!!!!!!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Write result to file
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function SaveChart() As String
Dim sRes As String, sPath As String

    On Error GoTo ErrHandle
'------------------------
If m_Res = "" Then Exit Function
If TC.fileName = "" Then TC.fileName = CreateOutputFileName
If m_Folder = "" Then m_Folder = CreateOutPutFolder
If m_Folder = "" Then Err.Raise 10008, , "Can't create  the folder"

sPath = m_Folder & "\" & TC.fileName: If Dir(sPath) <> "" Then Kill sPath
If WriteStringToFileUTF8(m_Res, sPath) Then sRes = sPath
'------------------------
ExitHere:
    SaveChart = sRes '!!!!!!!!
    Exit Function
'-------
ErrHandle:
    ErrPrint2 "SaveChart", Err.Number, Err.Description, MOD_NAME
    Err.Clear
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Create Config Chart
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Sub CreateConfig(ChartType As TCHARTTYPE)

    On Error Resume Next
'-------------------------
Select Case ChartType
    Case CHR_BAR:
          TC.ChartID = "CHRBAR"
          TC.ChartName = "BAR CHART"
          TC.DrawFuncName = "createBar"
          TC.TheClass = "BarChart"
          TC.ThePackageName = "corechart"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_BAR.html"
         
    Case CHR_BUBBLE:
          TC.ChartID = "CHRBUBBLE"
          TC.ChartName = "BUBBLE CHART"
          TC.DrawFuncName = "createBubble"
          TC.TheClass = "BubbleChart"
          TC.ThePackageName = "corechart"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_BUBBLE.html"
          
    Case CHR_COLUMNS:
          
          If Not Me.IsMaterial Then
                TC.TheClass = "ColumnChart"
                TC.ThePackageName = "corechart"
          Else
                TC.TheClass = "Bar"
                TC.ThePackageName = "bar"
          
          End If
                TC.ChartID = "CHRCOLUMN"
                TC.ChartName = "COLUMN CHART"
                TC.DrawFuncName = "createCOLUMNS"
                TC.APIProvider = GetLoader()
                TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_COLUMNS.html"
         
          
    Case CHR_DONUTS:
          TC.ChartID = "CHRDONUT"
          TC.ChartName = "DONUT CHART"
          TC.DrawFuncName = "createDonut"
          TC.TheClass = "PieChart"
          TC.ThePackageName = "corechart"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_DONUT.html"
          
          
    Case CHR_GANTT:
          TC.ChartID = "CHRGANTT"
          TC.ChartName = "GANTT CHART"
          TC.DrawFuncName = "createGANTT"
          TC.TheClass = "Gantt"
          TC.ThePackageName = "gantt"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_GANTT.html"
    
    
    Case CHR_GAUGE:
          TC.ChartID = "CHRGAUGE"
          TC.ChartName = "GAUGE CHART"
          TC.DrawFuncName = "createGauge"
          TC.TheClass = "Gauge"
          TC.ThePackageName = "gauge"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_GAUGE.html"
    
    Case CHR_GEO:
    
          TC.ChartID = "CHRGEO"
          TC.ChartName = "GEO CHART"
          TC.DrawFuncName = "createGeo"
          TC.TheClass = "GeoChart"
          TC.ThePackageName = "geochart"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_GEO.html"
    
    
    Case CHR_LINE:
          TC.ChartID = "CHRLINE"
          TC.ChartName = "LINE CHART"
          TC.DrawFuncName = "createLINE"
          TC.TheClass = "LineChart"
          TC.ThePackageName = "corechart"
          TC.APIProvider = GetLoader()
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_LINE.html"
    Case CHR_PIE:
          TC.ChartID = "CHRPIE"
          TC.ChartName = "PIE CHART"
          TC.DrawFuncName = "createPIE"
          TC.TheClass = "PieChart"
          TC.ThePackageName = "corechart"
          TC.fileName = "CHR_BAR"
          TC.APIProvider = GetJSAPI(0)
          TC.fileName = Format(Now(), "yyyymmdd") & "_" & "CHR_PIE.html"
    Case CHR_TIMELINES:
End Select

End Sub
'------------------------------------------------------------------------------------------------------------------------------------------------------
' Special replacement for field list
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function GetFieldReplacement(sFldReplName As String) As String
Dim sRes As String
    
    On Error Resume Next
'-------------------
    Select Case sFldReplName
    Case "ROLE_STYLE":
         sRes = "{ role: " & sCH("style") & " }"
    Case "ROLE_ANNOTATION":
         sRes = "{ role: " & sCH("annotation") & " }"
    Case "ROLE_TOOLTIP":
         sRes = "{ role: " & sCH("tooltip") & " }"
    Case "ROLE_INTERVAL":
         sRes = "{ role: " & sCH("interval") & " }"
    Case "ROLE_CERTAINTY":
         sRes = "{ role: " & sCH("certainty") & " }"
    Case "ROLE_SCOPE":
         sRes = "{ role: " & sCH("scope") & " }"
    Case Else
         sRes = sFldReplName
    End Select
'-------------------
ExitHere:
    GetFieldReplacement = sRes '!!!!!!!!
End Function

'------------------------------------------------------------------------------------------------------------------------------------------------------
' Date converter
'------------------------------------------------------------------------------------------------------------------------------------------------------
Private Function ISO8601Date(sDate As String) As String

If sDate = "" Then Exit Function
ISO8601Date = Format(sDate, "yyyy-mm-dd") '!!!!!!!!
End Function
