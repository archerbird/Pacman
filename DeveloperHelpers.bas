Attribute VB_Name = "DeveloperHelpers"
'@Folder "PacmanGame.Common"
Option Explicit

Sub ChangeWallTokens()
    Dim rng As Range
    Set rng = Selection
    Dim cell As Range
    
    For Each cell In rng
        If cell.value = "*" Then
            cell.Font.Color = vbBlack
        End If
    Next
End Sub

Function RecordMapEncoding() As String
    Dim output As String
    Dim outputArr(1 To 53) As String
    Dim outerOutputArr(1 To 59) As String
    Dim i As Integer
    Dim j As Integer
    Dim rep As Range
    Dim rRow As Range
    Dim cell As Range
    Set rep = Range("$D$3:$BD$61")
    
    For Each rRow In rep.Rows
        j = j + 1
        i = 0
        For Each cell In rRow.Cells
            i = i + 1
            If cell.value = "*" Then
                outputArr(i) = "w"  '//wall
            ElseIf cell.value = "•" Then
                If cell.Font.Size > 6 Then
                    outputArr(i) = "P" '//super ellet
                Else
                    outputArr(i) = "p" '//pellet
                End If
            ElseIf cell.value = "`" Then
                 outputArr(i) = "D"
            ElseIf cell.value = "~" Then
                outputArr(i) = "s"
            ElseIf cell.value = "d" Then
                outputArr(i) = "d"
            Else
                outputArr(i) = "m" '//regular maze pathing
            End If
        Next
        outerOutputArr(j) = Join(outputArr, ",")
    Next
    
    RecordMapEncoding = Join(outerOutputArr, ";")
End Function



'Sub ToTheClipboard(Text As String)
'Dim MyDataObj As New DataObject
'MyDataObj.SetText Text
'MyDataObj.PutInClipboard
'End Sub
'
'Sub PastingFromTheClipboard()
'Dim MyDataObj As New DataObject
'MyDataObj.GetFromClipboard
'
'Dim MyVar As Variant
'MyVar = MyDataObj.GetText
'MsgBox MyVar
'End Sub

Sub RecordGhostShape()
    Dim sg As GroupShapes
    Dim s As Shape
    Dim ghostShape As Shape
    
    Set ghostShape = Sheet1.Shapes("Ghost")
    Set sg = ghostShape.GroupItems
    For Each s In sg
        s.Select
        
        Debug.Print "{"
        Debug.Print "Type: " & s.AutoShapeType
        Debug.Print "Left: " & s.Left
        Debug.Print "Top: " & s.Top
        Debug.Print "Width: " & s.Width
        Debug.Print "Height: " & s.Height
        Debug.Print "Rotation: " & s.Rotation
        Debug.Print "HFlip: " & s.HorizontalFlip
        Debug.Print "VFlip: " & s.VerticalFlip
        Debug.Print "adjustments: " & s.Adjustments(1)
        Debug.Print "             " & s.Adjustments(2)
        Debug.Print s.Line.Visible
        Debug.Print "}"
        Debug.Print "--------------"
        
    Next
    
End Sub

Sub test()
Dim g As New GhostStyler

g.Init ActiveSheet, vbCyan

End Sub
