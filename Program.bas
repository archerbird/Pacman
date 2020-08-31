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
    
    '//load the map
    mController.Maze = DeveloperHelpers.LoadMapFromFile
    
    '//start the game!
    mController.StartGame viewUIAdapter
End Sub

Public Sub Quit()
    Set mController = Nothing
End Sub

