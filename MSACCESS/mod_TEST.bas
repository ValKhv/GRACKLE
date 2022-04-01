Attribute VB_Name = "mod_TEST"
Option Compare Database
Option Explicit


Private Const MOD_NAME As String = "mod_TEST"
'######################

Public Sub Test_GoogleCharts()
Dim MYCHART As New cGoogleCharts, sField As String
Dim sRes As String

    'sField = MYCHART.FieldName(0)
    'Debug.Print sField
    
    'MYCHART.Item(0, sField) = "2015"
    'Debug.Print MYCHART.Item(0, sField)
    ' sRes = MYCHART.Demo_CHART_PIE()
    'sRes = MYCHART.Demo_CHART_LINE
    'sRes = MYCHART.Demo_CHART_DONUTS
    'sRes = MYCHART.Demo_CHART_BAR_1
    'sRes = MYCHART.Demo_CHART_BAR_2
    'sRes = MYCHART.Demo_CHART_COLUMNS_1
    'sRes = MYCHART.Demo_CHART_COLUMNS_2
    'sRes = MYCHART.Demo_CHART_COLUMNS_3
    'sRes = MYCHART.Demo_CHART_GAUGE
    'sRes = MYCHART.Demo_CHART_BUBBLE_1
    'sRes = MYCHART.Demo_CHART_BUBBLE_2
    'sRes = MYCHART.Demo_CHART_GANTT_1
    'sRes = MYCHART.Demo_CHART_GANTT_2
    sRes = MYCHART.Demo_CHART_GEO
    
    Debug.Print sRes

Set MYCHART = Nothing
End Sub





