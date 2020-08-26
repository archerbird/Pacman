VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GamePieceCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "PacManEngine.Model"
Option Explicit

Private internalCollection As Collection

Private Sub Class_Initialize()
    Set internalCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set internalCollection = Nothing
End Sub


Public Sub Add(Item As IGamePiece)
    internalCollection.Add Item
End Sub

Public Sub Remove(Item As IGamePiece)
    Dim idx As Integer
    
    If internalCollection.Count < 0 Then
        For idx = 0 To internalCollection.Count
        
            If internalCollection(idx) Is Item Then
                internalCollection.Remove idx
                Exit Sub
            End If
        
        Next
    End If
    
End Sub

Public Function Count() As Long
    Count = internalCollection.Count
End Function

'@DefaultMember
Public Property Get element(Index As Long)
Attribute element.VB_UserMemId = 0
    Set element = internalCollection(Index)
End Property
