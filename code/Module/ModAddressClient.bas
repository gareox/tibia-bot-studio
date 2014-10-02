Attribute VB_Name = "ModAddressClient"
 Option Explicit
 
Public adrXOR As Long  ' &H7BF1F0
Public IsConeted As Long  ' &H7C8FF8

Public adrNChar As Long  ' &H953008
Public Player_ID As Long

Public Player_HP As Long  ' &H953000
Public Player_MP As Long  ' &H7BF244

Public AdrXPos As Long  ' &H98AEA8
Public AdrYPos As Long  ' AdrXPos + 4
Public AdrZPos As Long  ' AdrXPos + 8

Public GoToX As Long  ' &H98AEA0
Public GoToY As Long  ' GoToX - 8
Public GoToZ As Long  ' GoToX + 16

Public adrGo As Long  ' &H953058

Public RedSquare As Long  ' Experience + &H40
Public GreenSquare As Long  ' RedSquare - 4
Public WhiteSquare As Long  ' GreenSquare - 8
Public CharDist As Long

Public LAST_BATTLELISTPOS As Long
