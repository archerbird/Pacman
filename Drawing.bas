Attribute VB_Name = "Drawing"
'@Folder("PacManEngine")
Option Explicit

Private Const TEXT_DOT_CHAR_CODE As Integer = 149

Public Const clBLUE As Long = &HFF0000

Public Property Get TextDot() As String
    TextDot = Chr(TEXT_DOT_CHAR_CODE)
End Property

