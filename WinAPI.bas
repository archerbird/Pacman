Attribute VB_Name = "WinAPI"
'@Folder("PacManEngine.UI.Implementations.ExcelWorksheet")
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#Else
    Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
#End If



