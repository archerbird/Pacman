Attribute VB_Name = "Module1"
'@Folder("PacManEngine.GridCells")
Option Explicit

Public Enum Directional
    dUp
    dDown
    dLeft
    dRight
End Enum

Public Function RangeDistance(rng1 As Range, rng2 As Range) As Double
    Dim a As Long
    Dim b As Long
    
    a = (rng1.Row - rng2.Row) ^ 2
    b = (rng1.Column - rng2.Column) ^ 2
    RangeDistance = Sqr(a + b)
End Function

