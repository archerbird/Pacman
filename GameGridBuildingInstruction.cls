VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GameGridBuildingInstruction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBAProject")
Option Explicit

'//Always start in top left corner of map
'//q = UpperLeftCornerWall
'//a = LeftWall
'//z = LowerLeftCornerWall
'//w = UpperWall
'//x = LowerWall
'//e = UpperRightCornerWall
'//d = RightWall
'//c = LowerRightCornerWall
'//t = UpperRightOutsideCornerWall
'//r = UpperLeftOusideCornerWall
'//g = LowerRightOutsideConerWall
'//f = LowerLeftOutsideCornerWall

'//[ = LeftDoor
'//u = UpperDoor
'//n = LowerDoor
'//] = RightDoor

'//UpperCaseLetter means outside wall
'//LowerCaseLetter means inside wall

'//_ = NoWallNoPellet
'//+ = HasPellet
'//* = HasSuperPellet

'//P = PacMan Starting Location
'//b = Blinky Start Location
'//p = Pinky Start Location
'//i = Inky Start Location
'//k = Clyde start Location


Private Const DEFAULT_MAP As String = "Q23WE3 Q23WE;A+_+_+_+_+_+_+_+_+_+_+_+D3 A+_+_+_+_+_+_+_+_+_+_+_+D;A_r7xt_r9xt_D3 A_r9xt_r7xt_D;"


