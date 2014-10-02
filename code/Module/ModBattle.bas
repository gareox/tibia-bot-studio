Attribute VB_Name = "ModBattle"
 Option Explicit

Public StepCreatures As Long ' = &HA0
Public MaxCreatures As Long '= 187

Public BattleList_Start As Long  ' = &H953008
Public Distance_Characters As Long '= 176

Public BattleList_End As Long  ' = BattleList_Start + (StepCreatures * MaxCreatures)

Public BattleList_Address As Long

Public Blatte_First As Long  '= BattleList_Start - 4

Public Enum distance
  cID = 0
  cType = 3
  Name = 4
  x = 44
  y = 40
  Z = 36
  IsWalking = 80
  Direction = 84
  Outfit = 96
  Addon = 116
  MountId = 120
  IsVisible = 172
  Skull = 152
  ColorHead = 100
  ColorBody = 104
  ColorLegs = 108
  ColorFeet = 112
  WarIcon = 168
  WalkSpeed = 144
  HPBar = 140
  Party = 152
End Enum

Public tMonters() As String
