Attribute VB_Name = "Client"
'@Folder "ViewImplementations.ExcelWorksheet"
Option Explicit
Private mController As GameController



Public Sub Main()
    'set up the UI
    
    'set up the game controller

End Sub

Public Sub Prototype()
    '//get our concrete sheet
    Dim xlWs As Worksheet
    Set xlWs = Sheet1
    
    '//wrap it up
    Dim sheetWrapper As WorksheetViewWrapper
    Set sheetWrapper = New WorksheetViewWrapper
    sheetWrapper.Init xlWs

    '//give it to a game adapter
    Dim viewUIAdapter As ViewAdapter
    Set viewUIAdapter = New ViewAdapter
    viewUIAdapter.Init sheetWrapper
    
    '//hand that to a new controller
    Set mController = New GameController
    Set mController.UIAdapter = viewUIAdapter

    '//start the game!
    mController.StartGame
End Sub

Public Sub Quit()
    Set mController = Nothing
End Sub

