Attribute VB_Name = "Program"
'@Folder("PacManEngine")
Option Explicit
Private mController As GameController


Public Sub Main()
    'set up the UI
    
    'set up the game controller

End Sub

Public Sub TestWithHardCodedSheet()
    '//get our concrete sheet
    Dim xlWs As Worksheet
    Set xlWs = PacmanUI
    
    '//wrap it up
    Dim sheetWrapper As WorksheetUIWrapper
    Set sheetWrapper = New WorksheetUIWrapper
    sheetWrapper.Init xlWs

    '//hand it to an Excel Adapter
    Dim xlUIAdapter As ExcelUIAdapter
    Set xlUIAdapter = New ExcelUIAdapter
    xlUIAdapter.Init sheetWrapper
    
    '//give the exce adapter to a game adapter
    Dim viewUIAdapter As GameUIAdapter
    Set viewUIAdapter = New GameUIAdapter
    viewUIAdapter.Init xlUIAdapter
    
    '//hand that to a new controller
    Set mController = New GameController
    
    '//load the map
    mController.Maze = DeveloperHelpers.LoadMapFromFile
    
    '//start the game!
    mController.StartGame viewUIAdapter
End Sub

Public Sub Quit()
    Set mController = Nothing
End Sub

