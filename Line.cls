VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Line"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//Configuation items
Public Color As Long
Public Style As XlLineStyle
Public weight As XlBorderWeight


Public Sub Draw(rangeBorder As Border)

   With rangeBorder
        .Color = Me.Color
        .lineStyle = Me.Style
        .weight = Me.weight
    End With
   
End Sub
