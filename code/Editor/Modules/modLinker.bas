Attribute VB_Name = "modLinker"
Option Explicit

Enum ApplicationType
    Console = 3
    GUI = 2
End Enum

Public Passes As Integer
Public LinkCode() As Byte
Public SectionSize As Long
Public AppType As ApplicationType

Sub InitLinker()
    ReDim LinkCode(0) As Byte
    InitDOSHeader
    InitDOSStub
    InitPEHeader
    InitSectionTable
    InitRawData
    InitCodeSection
    InitImportSection
End Sub

Function CheckSectionSize() As Boolean
    Dim i As Integer
    Dim cSize As Long
    Dim Value As Long
    
 
    Value = UBound(CodeSection)
    If Value < UBound(DataSection) Then Value = UBound(DataSection)
    If Value < UBound(ImportSection) Then Value = UBound(ImportSection)
    
    For i = 0 To Value Step 512
        cSize = cSize + 2
    Next i
    
    If SectionSize < cSize Then
        SectionSize = cSize
        Passes = Passes + 1
        CheckSectionSize = True
        Exit Function
    End If
    
    CheckSectionSize = False
    
End Function

Sub InitDOSHeader()
    AddLinkByte &H4D, &H5A, &H80, &H0, &H1, &H0, &H0, &H0, &H4, &H0, &H10, &H0, &HFF, &HFF, &H0, &H0
    AddLinkByte &H40, &H1, &H0, &H0, &H0, &H0, &H0, &H0, &H40, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H80, &H0, &H0, &H0
End Sub

Sub InitDOSStub()
    AddLinkByte &HE, &H1F, &HBA, &HE, &H0, &HB4, &H9, &HCD, &H21, &HB8, &H1, &H4C, &HCD, &H21, &H74
    AddLinkByte &H68, &H69, &H73, &H20, &H70, &H72, &H6F, &H67, &H72, &H61, &H6D, &H20, &H63, &H61
    AddLinkByte &H6E, &H6E, &H6F, &H74, &H20, &H62, &H65, &H20, &H72, &H75, &H6E, &H20, &H69, &H6E
    AddLinkByte &H20, &H44, &H4F, &H53, &H20, &H6D, &H6F, &H64, &H65, &H2E, &HD, &HA, &H24, &H0, &H0
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0
End Sub

Sub InitPEHeader()
    'Signature = "PE"
    AddLinkByte &H50, &H45, &H0, &H0
    
    'Machine 0x014C;i386
    AddLinkByte &H4C, &H1
    'NumberOfSections = 3
    AddLinkByte &H3, &H0
    'TimeDateStamp
    AddLinkByte &H0, &H0, &H0, &H0
    'PointerToSymbolTable = 0
    AddLinkByte &H0, &H0, &H0, &H0
    'NumberOfSymbols = 0
    AddLinkByte &H0, &H0, &H0, &H0
    'SizeOfOptionalHeader
    AddLinkByte &HE0, &H0
    'Characteristics
    AddLinkByte &H8F, &H81
    
    'Magic
    AddLinkByte &HB, &H1
    'MajorLinkerVersion
    AddLinkByte &H1
    'MinerLinkerVersion
    AddLinkByte &H0
    'SizeOfCode
    AddLinkByte &H0, &H0, &H0, &H0
    'SizeOfInitializedData
    AddLinkByte &H0, &H0, &H0, &H0
    'SizeOfUnInitializedData
    AddLinkByte &H0, &H0, &H0, &H0
    'AddressOfEntryPoint
    AddLinkByte &H0, &H20, &H0, &H0
    'BaseOfCode
    AddLinkByte &H0, &H0, &H0, &H0
    'BaseOfData
    AddLinkByte &H0, &H0, &H0, &H0
    'ImageBase
    AddLinkByte &H0, &H0, &H40, &H0
    'SectionAlignment
    AddLinkByte &H0, &H10, &H0, &H0
    'FileAlignment
    AddLinkByte &H0, &H2, &H0, &H0
    'MajorOSVersion
    AddLinkByte &H1, &H0
    'MinorOSVersion
    AddLinkByte &H0, &H0
    'MajorImageVersion
    AddLinkByte &H0, &H0
    'MinorImageVersion
    AddLinkByte &H0, &H0
    'MajorSubSystemVerion
    AddLinkByte &H4, &H0
    'MinorSubSystemVerion
    AddLinkByte &H0, &H0
    'Win32VersionValue
    AddLinkByte &H0, &H0, &H0, &H0
    'SizeOfImage
    AddLinkByte &H0, &H40, &H0, &H0
    'SizeOfHeaders
    AddLinkByte &H0, &H2, &H0, &H0
    'CheckSum
    AddLinkByte &H94, &H6B, &H0, &H0
    'SubSystem = 2:GUI; 3:Console
    AddLinkByte AppType
    AddLinkByte &H0
    'DllCharacteristics
    AddLinkByte &H0, &H0
    'SizeOfStackReserve
    AddLinkByte &H0, &H10, &H0, &H0
    'SizeOfStackCommit
    AddLinkByte &H0, &H10, &H0, &H0
    'SizeOfHeapReserve
    AddLinkByte &H0, &H0, &H1, &H0
    'SizeOfHeapRCommit
    AddLinkByte &H0, &H0, &H0, &H0
    'LoaderFlags
    AddLinkByte &H0, &H0, &H0, &H0
    'NumberOfDataDirectories
    AddLinkByte &H10, &H0, &H0, &H0
    
    'Export_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Import_Table
    AddLinkByte &H0, &H30, &H0, &H0, &H80, &H0, &H0, &H0
    'ReSource_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Exception_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Certificate_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Relocation_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Debug_Data
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Architecture
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Global_PTR
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'TLS_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Load_Config_Table
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'BoundImportTable
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'ImportAddressTable
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'DelayImportDescriptor
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'COMplusRuntimeHeader
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
    'Reserved
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
End Sub

Sub InitSectionTable()
    'Section1
    'Name = ".data"
    AddLinkByte &H2E, &H64, &H61, &H74, &H61, &H0, &H0, &H0
    'VirtualSize
    'AddLinkByte &H33, &H0, &H0, &H0
    SizeOfRawData 1
    'VirtualAddress
    AddLinkByte &H0, &H10, &H0, &H0
    
        'SizeOfRawData
        SizeOfRawData 1
        'PointerToRawData
        PointerToRawData 1, SectionSize
    
    'PointerToRelocations
    AddLinkByte &H0, &H0, &H0, &H0
    'PointerToLinenumbers
    AddLinkByte &H0, &H0, &H0, &H0
    'NumberOfRelocations
    AddLinkByte &H0, &H0
    'NumberOfLinenumbers
    AddLinkByte &H0, &H0
    'Characteristics
    AddLinkByte &H40, &H0, &H0, &HC0
    
    'Section2
    'Name = ".code"
    AddLinkByte &H2E, &H63, &H6F, &H64, &H65, &H0, &H0, &H0
    'VirtualSize
    SizeOfRawData 2
    'AddLinkByte &H1C, &H0, &H0, &H0
    'VirtualAddress
    AddLinkByte &H0, &H20, &H0, &H0
    
        'SizeOfRawData
        SizeOfRawData 2
        'PointerToRawData
        PointerToRawData 2, SectionSize
    
    'PointerToRelocations
    AddLinkByte &H0, &H0, &H0, &H0
    'PointerToLinenumbers
    AddLinkByte &H0, &H0, &H0, &H0
    'NumberOfRelocations
    AddLinkByte &H0, &H0
    'NumberOfLinenumbers
    AddLinkByte &H0, &H0
    'Characteristics
    AddLinkByte &H20, &H0, &H0, &H60
    
    'Section3
    'Name = ".idata"
    AddLinkByte &H2E, &H69, &H64, &H61, &H74, &H61, &H0, &H0
    'VirtualSize
    SizeOfRawData 3
    'AddLinkByte &H80, &H0, &H0, &H0
    'VirtualAddress
    AddLinkByte &H0, &H30, &H0, &H0
    
        'SizeOfRawData
        SizeOfRawData 3
        'PointerToRawData
        PointerToRawData 3, SectionSize
    
    'PointerToRelocations
    AddLinkByte &H0, &H0, &H0, &H0
    'PointerToLinenumbers
    AddLinkByte &H0, &H0, &H0, &H0
    'NumberOfRelocations
    AddLinkByte &H0, &H0
    'NumberOfLinenumbers
    AddLinkByte &H0, &H0
    'Characteristics
    AddLinkByte &H40, &H0, &H0, &HC0
    
    AddLinkByte &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0, &H0
End Sub

Sub InitRawData()
    Dim i As Long
    Dim SizeOfData As Long
    
    SizeOfData = UBound(DataSection)
    
    For i = SizeOfData To (256 * SectionSize) - 1
        AddDataByte &H0
    Next i
        
    For i = 1 To UBound(DataSection)
        AddLinkByte DataSection(i)
    Next i
End Sub

Sub InitCodeSection()
    Dim i As Long
    Dim SizeOfCode As Long
    
    SizeOfCode = UBound(CodeSection)
    
    For i = SizeOfCode To (256 * SectionSize) - 1
        AddCodeByte &H0
    Next i
        
    For i = 1 To UBound(CodeSection)
        AddLinkByte CodeSection(i)
    Next i
End Sub

Sub InitImportSection()
    Dim i As Long
    Dim SizeOfImport As Long
    
    OutputImportTable
    
    SizeOfImport = UBound(ImportSection)
    
    For i = SizeOfImport To (256 * SectionSize) - 1
        AddImportByte &H0
    Next i
        
    For i = 1 To UBound(ImportSection)
        AddLinkByte ImportSection(i)
    Next i
End Sub

Sub SizeOfRawData(Section As Long)
    AddLinkDWord CLng(("&H" & Hex(SectionSize))) * &H100
End Sub

Sub PointerToRawData(Section As Integer, SizeB As Long)
    Dim i As Integer
    Dim Value As Long
    Value = 2
    For i = 2 To Section
        Value = Value + SizeB
    Next i
    AddLinkDWord CLng("&H" & Hex(Value * &H100))
End Sub

Sub Link(sFile As String, Run As Boolean)
    Dim i As Double
    
    If AppType = 0 Then
        InfMessage "application type not specified. compiling process failed."
        Exit Sub
    End If
    
    On Error GoTo LinkFail
    If Dir(sFile) <> "" Then Kill sFile
    
    Open sFile For Binary As #1
        For i = 1 To UBound(LinkCode)
            Put #1, , LinkCode(i)
        Next i
    Close #1
    InfMessage "application compiled. " & vbCrLf & _
               Passes & " pass(es). " & UBound(LinkCode) & " bytes written."
    
    If Run = True Then ShellExecute 0, "open", sFile, "", "C:\", 1
    Exit Sub
LinkFail:
    InfMessage "linking process failed. unknown reason."
End Sub

Sub AddLinkDWord(Value As Long)
    AddLinkWord LoWord(Value)
    AddLinkWord HiWord(Value)
End Sub

Sub AddLinkWord(Value As Integer)
    AddLinkByte LoByte(Value), HiByte(Value)
End Sub

Sub AddLinkByte(ParamArray Bytes() As Variant)
    Dim i As Integer
    For i = 0 To UBound(Bytes)
        ReDim Preserve LinkCode(UBound(LinkCode) + 1) As Byte
        LinkCode(UBound(LinkCode)) = Bytes(i)
    Next i
End Sub

