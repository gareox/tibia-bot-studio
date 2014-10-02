Attribute VB_Name = "ModFuncions"
 Option Explicit

Public Function PlayerOnline() As Long
    PlayerOnline = Memory_ReadLong(Tibia_Hwnd, theOffset + IsConeted)
End Function

Public Function PlayerID() As Long
    PlayerID = Memory_ReadLong(Tibia_Hwnd, theOffset + adrNChar)
End Function

Public Function PlayerHP() As Long
    Dim tmp As Long
    Dim valueXOR As Long
    
    tmp = Memory_ReadLong(Tibia_Hwnd, theOffset + Player_HP)
    valueXOR = Memory_ReadLong(Tibia_Hwnd, theOffset + adrXOR)
    tmp = valueXOR Xor tmp
    PlayerHP = tmp
End Function

Public Function PlayerMP() As Long
    Dim tmp As Long
    Dim valueXOR As Long
    
    tmp = Memory_ReadLong(Tibia_Hwnd, theOffset + Player_MP)
    valueXOR = Memory_ReadLong(Tibia_Hwnd, theOffset + adrXOR)
    tmp = valueXOR Xor tmp
    PlayerMP = tmp
End Function

Public Function PlayerAttacking() As Long
    PlayerAttacking = Memory_ReadLong(Tibia_Hwnd, theOffset + RedSquare)
End Function

Public Function PlayerX() As Long
    PlayerX = Memory_ReadLong(Tibia_Hwnd, theOffset + AdrXPos)
End Function

Public Function PlayerY() As Long
    PlayerY = Memory_ReadLong(Tibia_Hwnd, theOffset + AdrYPos)
End Function

Public Function PlayerZ() As Long
    PlayerZ = Memory_ReadLong(Tibia_Hwnd, theOffset + AdrZPos)
End Function

Public Function MyBattleListPosition() As Long
  Dim c1 As Long
  Dim ID As Double
  Dim res As Long
  On Error GoTo goterr
  res = -1
  For c1 = 0 To LAST_BATTLELISTPOS
    ID = CDbl(Memory_ReadLong(Tibia_Hwnd, theOffset + adrNChar + (CharDist * c1)))
    If myID = ID Then
      res = c1
      Exit For
    End If
  Next c1
  MyBattleListPosition = res
  Exit Function
goterr:
  MyBattleListPosition = -1
End Function

Public Function PlayerWalk(x As Long, y As Long, Z As Long)
    
    Call MemoryWriteLong(Tibia_Hwnd, theOffset + GoToX, x)
    Call MemoryWriteLong(Tibia_Hwnd, theOffset + GoToY, y)
    Call MemoryWriteLong(Tibia_Hwnd, theOffset + GoToZ, Z)

    Memory_WriteByte Tibia_Hwnd, theOffset + adrGo + (myBpos * CharDist), 1
End Function


Public Function IsPlayerWalking() As Integer
    IsPlayerWalking = Memory_ReadByte(Tibia_Hwnd, theOffset + adrGo + (myBpos * CharDist))
End Function

Public Sub Attacker(Name As String)
Dim cName As String
Dim cID As Long
Dim cx As Long, cy As Long, cZ As Long
Dim IsVisible As Long

If PlayerAttacking = 0 Then 'Si el personaje no esta atacando (Continuamos)

        For BattleList_Address = BattleList_Start To BattleList_End Step Distance_Characters
        DoEvents
        IsVisible = Memory_ReadLong(Tibia_Hwnd, theOffset + BattleList_Address + distance.IsVisible)
        cName = Memory_ReadString(Tibia_Hwnd, theOffset + BattleList_Address + distance.Name)
        cName = LCase(cName)
        
        If InStr(cName, LCase(Name)) = 1 Then
            
            cID = Memory_ReadLong(Tibia_Hwnd, theOffset + BattleList_Address + distance.cID)
        
            If IsVisible <> 0 Then
            
                cx = Memory_ReadLong(Tibia_Hwnd, theOffset + BattleList_Address + distance.x)
                cy = Memory_ReadLong(Tibia_Hwnd, theOffset + BattleList_Address + distance.y)
                cZ = Memory_ReadLong(Tibia_Hwnd, theOffset + BattleList_Address + distance.Z)
                
                If cZ = PlayerZ Then
                    If cx + 7 >= PlayerX And cx - 7 <= PlayerX Then
                        If cy + 5 >= PlayerY And cy - 5 <= PlayerY Then
                
                            Call Memory_WriteLong(Tibia_Hwnd, theOffset + RedSquare, cID) 'lo marca con un cuadro rojo.
                            PlayerAtack cID 'Atacamos
                            Exit For    'Salimos de la funcion
                        End If
                    End If
               End If
               
            End If
            
        End If
            DoEvents
    Next BattleList_Address
    
End If

End Sub

