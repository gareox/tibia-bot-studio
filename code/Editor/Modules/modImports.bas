Attribute VB_Name = "modImports"
Option Explicit

Type ITEM_IIBN
    ThunkValue As Long
    Hint As Integer
    Name As String
End Type

Type ITEM_ALIAS
    FunctionName As String
    FunctionAlias As String
    Library As String
    ParamCount As Long
    ImageImportByNameID As Long
End Type

Public TempDLLOffset As Long
Public TempNameOffset As Long

Public DLLNames() As String
Public ImageImportByName() As ITEM_IIBN
Public ImageImportDescriptor() As Byte

Public Aliases() As ITEM_ALIAS
Public iUniqueID As Long
Public ImportSection() As Byte

Sub InitImports()
    iUniqueID = 0
    ReDim Aliases(0) As ITEM_ALIAS
    ReDim DLLNames(0) As String
    ReDim ImportSection(0) As Byte
    ReDim ImageImportByName(0) As ITEM_IIBN
    ReDim ImageImportDescriptor(0) As Byte
End Sub

Sub AddAlias(Name As String, Library As String, ParamCount As Long, Optional Alias As String)
   
    ReDim Preserve Aliases(UBound(Aliases) + 1) As ITEM_ALIAS
    Aliases(UBound(Aliases)).FunctionName = Name
    If Alias <> "" Then
        Aliases(UBound(Aliases)).FunctionAlias = Alias
    Else
        Aliases(UBound(Aliases)).FunctionAlias = Name
    End If
    Aliases(UBound(Aliases)).Library = Library
    Aliases(UBound(Aliases)).ParamCount = ParamCount
    Aliases(UBound(Aliases)).ImageImportByNameID = UBound(ImageImportByName) + 1
    
    ReDim Preserve DLLNames(UBound(DLLNames) + 1) As String
    DLLNames(UBound(DLLNames)) = Library
    
    ReDim Preserve ImageImportByName(UBound(ImageImportByName) + 1) As ITEM_IIBN
    ImageImportByName(UBound(ImageImportByName)).Hint = &H0
    ImageImportByName(UBound(ImageImportByName)).Name = Name
    
    ReDim Preserve ImageImportDescriptor(UBound(ImageImportDescriptor) + 1) As Byte
End Sub

Sub OutputImportTable()
    Dim i As Integer
    Dim ii As Integer
    
    For i = 1 To UBound(ImageImportDescriptor)
        AddImportDWord &H0
        AddImportDWord &H0
        AddImportDWord &H0
        AddImportDWord &H0
        AddFixup "DLLName" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection), 0, Import
        AddImportDWord &H0
        AddFixup "ImageImportByName" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection), 0, Import
    Next i
        
        AddImportDWord &H0
        AddImportDWord &H0
        AddImportDWord &H0
        AddImportDWord &H0
        AddImportDWord &H0
        
    For i = 1 To UBound(DLLNames)
        AddSymbol "DLLName" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection), Import, IS_ALIAS
        For ii = 1 To Len(DLLNames(i))
            AddImportByte CByte(Asc(Mid(DLLNames(i), ii, 1)))
        Next ii
        AddImportByte &H0
    Next i
    
    For i = 1 To UBound(ImageImportByName)
        AddSymbol "ImageImportByName" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection), Import, IS_ALIAS
        AddImportDWord ImageImportByName(i).ThunkValue
        AddFixup "Hint" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection), 0, Import
        AddImportByte &H0
        AddSymbol "Hint" & i, (512 + (256 * SectionSize) + (256 * SectionSize)) + UBound(ImportSection) + 3, Import, IS_ALIAS
        AddImportWord ImageImportByName(i).Hint
        AddImportByte &H0, &H0, &H0
        For ii = 1 To Len(ImageImportByName(i).Name)
            AddImportByte CByte(Asc(Mid(ImageImportByName(i).Name, ii, 1)))
        Next ii
        AddImportByte &H0
    Next i

End Sub

Function GetImportIDByName(Name As String) As Long
    Dim i As Integer
    For i = 0 To UBound(Aliases)
        If Aliases(i).FunctionAlias = Name Then
            GetImportIDByName = Aliases(i).ImageImportByNameID
            Exit Function
        End If
    Next i
End Function

Function GetImportParameterCountByName(Name As String) As Long
    Dim i As Integer
    For i = 0 To UBound(Aliases)
        If Aliases(i).FunctionAlias = Name Then
            GetImportParameterCountByName = Aliases(i).ParamCount
        End If
    Next i
End Function

Sub AddImportDWord(ParamArray DWords() As Variant)
    Dim i As Integer
    For i = 0 To UBound(DWords)
        AddImportWord LoWord(CLng(DWords(i))), HiWord(CLng(DWords(i)))
    Next i
End Sub

Sub AddImportWord(ParamArray Words() As Variant)
    Dim i As Integer
    For i = 0 To UBound(Words)
        AddImportByte LoByte(CInt(Words(i))), HiByte(CInt(Words(i)))
    Next i
End Sub

Sub AddImportByte(ParamArray Bytes() As Variant)
    Dim i As Integer
    For i = 0 To UBound(Bytes)
        ReDim Preserve ImportSection(UBound(ImportSection) + 1) As Byte
        ImportSection(UBound(ImportSection)) = Bytes(i)
    Next i
End Sub
