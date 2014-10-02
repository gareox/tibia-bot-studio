Attribute VB_Name = "modFixups"
Option Explicit

Enum SECTION_SIZE
    Data = 3584
    Code = 10752
    Import = 10752
End Enum

Type ITEM_FIXUP
    Name As String
    Offset As Long
    Value As Long
    Section As SECTION_SIZE
End Type

Public SizeOffset As Long
Public Fixups() As ITEM_FIXUP

Sub InitFixups()
    ReDim Fixups(0) As ITEM_FIXUP
End Sub

Sub AddCodeFixup(Name As String)
    AddCodeDWord &H0
    AddFixup Name, (512 + (256 * SectionSize)) + UBound(CodeSection), &H400000, Data
End Sub

Sub AddFixup(Name As String, Offset As Long, Value As Long, Section As SECTION_SIZE)
    ReDim Preserve Fixups(UBound(Fixups) + 1) As ITEM_FIXUP
    Fixups(UBound(Fixups)).Name = Name
    Fixups(UBound(Fixups)).Offset = Offset
    Fixups(UBound(Fixups)).Value = Value
    Fixups(UBound(Fixups)).Section = Section
End Sub

Sub DoFixups()
    Dim i As Integer
    Dim ii As Integer
    Dim Found As Boolean
    Dim SetOffset(1) As Long
    
    For i = 1 To UBound(Fixups)
        For ii = 1 To UBound(Symbols)
            If Fixups(i).Name = Symbols(ii).Name Then
                If Symbols(ii).SymType = IS_LABEL Or _
                   Symbols(ii).SymType = IS_LOCAL_DWORD Or _
                   Symbols(ii).SymType = IS_LOCAL_STRING Or _
                   Symbols(ii).SymType = IS_FRAME Then
                    FixDWord Fixups(i).Offset - 3, Symbols(ii).Offset - Fixups(i).Value
                Else
                    If Fixups(i).Section = Data Then
                        FixDWord Fixups(i).Offset - 3, Symbols(ii).Offset + Fixups(i).Section + Fixups(i).Value
                    ElseIf Fixups(i).Section = Code Then
                        SizeOffset = (512 * SectionSize) - 1024
                        FixDWord Fixups(i).Offset - 3, Symbols(ii).Offset + (Fixups(i).Section - SizeOffset) + Fixups(i).Value
                    End If
                End If
                Found = True
            End If
        Next ii
        If Found = False Then ErrMessage "could not find " & Fixups(i).Name
        Found = False
    Next i
End Sub

