VERSION 5.00
Begin VB.UserControl McImageList 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1365
   FillColor       =   &H00FECFA5&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   91
   ToolboxBitmap   =   "McImageList.ctx":0000
End
Attribute VB_Name = "McImageList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^Gtech^^Creations^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^¶¶^^^^^¶¶^^^^^^^¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^¶¶^^^^^^^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^^$
'$^^¶¶¶^^^¶¶¶^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^¶¶^^^^^^^¶¶¶^^^^^^^¶¶¶^^^^$
'$^^¶¶¶¶^¶¶¶¶^^¶¶¶¶^^¶¶^^¶¶¶¶¶^¶¶¶^^^¶¶¶¶^^^¶¶¶¶¶^^¶¶¶¶^^¶¶^^^^¶¶^^¶¶¶¶^¶¶¶¶^^^^^^¶¶^^^^^^^^¶¶^^^^$
'$^^¶^¶¶¶¶^¶¶^¶¶^^^^^¶¶^^¶¶^^¶¶^^¶¶^^^^^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^^^¶¶^¶¶^^^^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^^$
'$^^¶^^¶¶^^¶¶^¶¶^^^^^¶¶^^¶¶^^¶¶^^¶¶^^¶¶¶¶¶^¶¶^^¶¶^¶¶¶¶¶¶^¶¶^^^^¶¶^¶¶¶¶^^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^^$
'$^^¶^^^^^^¶¶^¶¶^^^^^¶¶^^¶¶^^¶¶^^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^^^^¶¶^^^^¶¶^^¶¶¶¶^¶¶^^^^^^^^¶¶^^^^^^^^¶¶^^^^$
'$^^¶^^^^^^¶¶^¶¶^^^^^¶¶^^¶¶^^¶¶^^¶¶^¶¶^^¶¶^¶¶^^¶¶^¶¶^^^^^¶¶^^^^¶¶^^^^¶¶^¶¶^^^^^^^^¶¶^^^¶¶^^^¶¶^^^^$
'$^^¶^^^^^^¶¶^^¶¶¶¶^¶¶¶¶^¶¶^^¶¶^^¶¶^^¶¶¶¶¶^^¶¶¶¶¶^^¶¶¶¶¶^¶¶¶¶¶^¶¶^¶¶¶¶^^^¶¶¶^^^^^¶¶¶¶^^¶¶^^¶¶¶¶^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^¶¶¶¶^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^By^Jim^^Jose^^^^^^Email^jimjosev33@yahoo.com^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$

'------------------------------------------------------------------------------------------
' SourceCode : McImageList 1.2
' Auther     : Jim Jose
' Email      : jimjosev33@yahoo.com
' Date       : 2-9-2005
' Purpose    : A higly flexible, lightweight ImageList control
' CopyRight  : JimJose © Gtech Creations - 2005
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' About :
'
'       I was working with a ListBox Control where I have to set diiferent
' Icons for each items on list. But declaring 'picture objects' to each item
' is something crazy! (because it will cause heavy use of memory)
'
'       The only one alternative was to use an 'Imagelist' and make it an image
' referance to the 'ListBox'. But I don't want to include a heavy Ocx (MSCommonControls-6)
' So I thought about a new custom 'Imagelist' control.
'
'       But, to ADD/REMOVE/MODFY images we need to use a 'Property page', which
' will make the control heavy. There comes the new technique to use the vb's 'property
' window' to do all these..
'
'       Now the control is 'SingleFile', 'LightWeight' and highly 'Flexible'.
' All the image operations(adding,removing,...) are very simple!! and give
' a better appearance!!
'
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' Usage :
'
' ADD Image: Slect the property 'AddNewImage' from property window and
' browse for the image u need.
'
' REMOVE Image : Set the property 'CurrentImage' and chage the property
' value of 'RemoveImage' to '[Yes!]'
'
' MODIFY Image : Set the 'CurrentImage' and browse for a new image from the
' property 'Image'
'
' MOVE Images : Set the 'Currentimage' and enter the index of new position to
' the property 'MoveImageTo'
'
' Accessing Image : The property 'ListImages' will return the 'picture object'
' of an index
'
'------------------------------------------------------------------------------------------

'------------------------------------------------------------------------------------------
' Updates :
'   3/9/2005 - Removed the bug, noticed by LiTe | Thanks to LiTe
'   3/9/2005 - Added code to 'Erase' array for clearing the memmory
'   7/9/2005 - Preventing negative numbers from entering 'CurrImage' property is
'              prevented. Thanks to 'Heriberto' for informing the severe bug.
'   9/9/2005 - Some good fixup suggested by 'Roger Gilcrist'. Also added
'              one function 'CyclePic' ( by Roger ) for sequential use or animation purpose
'------------------------------------------------------------------------------------------

Option Explicit

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Enum RemoveItemEnum
    [No!] = 0
    [Yes!] = 1
End Enum

Private m_ImageCount    As Long
Private m_CurrentImage  As Long
Private m_Image()       As StdPicture
Private Const m_def_CurrentImage = 0

Public Property Get ListImages(ByVal vIndex As Long) As StdPicture
On Error Resume Next
    Set ListImages = m_Image(vIndex)
End Property

Public Property Set ListImages(ByVal vIndex As Long, ByVal vNewValue As StdPicture)
    Set m_Image(vIndex) = vNewValue
    RedrawControl
End Property

Public Property Get ImageCount() As Long
On Error GoTo handle
    ImageCount = UBound(m_Image) + 1
Exit Sub
handle:
ImageCount = 0
End Property

Public Property Let ImageCount(ByVal vNewValue As Long)

End Property

Public Property Get Image() As StdPicture
    On Error Resume Next
    Set Image = m_Image(m_CurrentImage)
End Property

Public Property Set Image(ByVal vNewValue As StdPicture)
    Set m_Image(m_CurrentImage) = vNewValue
    PropertyChanged "Image"
    RedrawControl
End Property

Public Property Get CurrentImage() As Long
    CurrentImage = m_CurrentImage
End Property

Public Property Let CurrentImage(ByVal New_CurrentImage As Long)
    If New_CurrentImage < 0 Then New_CurrentImage = 0
    m_CurrentImage = New_CurrentImage
    If m_CurrentImage > Me.ImageCount - 1 Then m_CurrentImage = Me.ImageCount - 1
    PropertyChanged "CurrentImage"
    RedrawControl
End Property

Private Sub UserControl_InitProperties()
    m_CurrentImage = m_def_CurrentImage
End Sub

Private Sub UserControl_Resize()
    RedrawControl
End Sub

Private Sub UserControl_Show()
    RedrawControl
End Sub

'Write ImageList
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim X As Long

    m_ImageCount = Me.ImageCount
    If m_ImageCount = 0 Then GoTo Skip
    For X = 0 To m_ImageCount - 1
        PropBag.WriteProperty "Images" & X, m_Image(X)
    Next X
    
Skip:
    PropBag.WriteProperty "ImageCount", m_ImageCount
    Call PropBag.WriteProperty("CurrentImage", m_CurrentImage, m_def_CurrentImage)
End Sub

' Read ImageList
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim X As Long
    
    On Error GoTo handle
    m_ImageCount = PropBag.ReadProperty("ImageCount")
    If m_ImageCount = 0 Then GoTo Skip
    ReDim m_Image(m_ImageCount - 1)
    For X = 0 To m_ImageCount - 1
        Set m_Image(X) = PropBag.ReadProperty("Images" & X)
    Next X
    
Skip:
    m_CurrentImage = PropBag.ReadProperty("CurrentImage", m_def_CurrentImage)

handle:
End Sub

'Property to add Item
Public Property Get AddNewImage() As StdPicture
End Property

Public Property Set AddNewImage(ByVal vNewValue As StdPicture)
Dim mArray() As StdPicture
Dim mCount As Long
Dim X As Long

    mCount = ImageCount
    If mCount = 0 Then
        ReDim m_Image(0)
        Set m_Image(0) = vNewValue
        
    Else
        mArray = m_Image
        Erase m_Image
        ReDim m_Image(mCount)
        For X = 0 To mCount - 1
            Set m_Image(X) = mArray(X)
        Next X
        Set m_Image(mCount) = vNewValue
        
    End If
    
    m_CurrentImage = ImageCount - 1
    RedrawControl

End Property


'Property to Remove Item
Public Property Get RemoveImage() As RemoveItemEnum
    RemoveImage = 0
End Property

Public Property Let RemoveImage(ByVal vNewValue As RemoveItemEnum)
Dim mArray() As StdPicture
Dim mCount As Long
Dim X As Long
Dim y As Long

    On Error GoTo handle
    If vNewValue = 0 Then Exit Property
    mCount = ImageCount
    If mCount = 0 Then Exit Property
    
    mArray = m_Image
    Erase m_Image
    ReDim m_Image(mCount - 2)
    
    For X = 0 To mCount - 1
        If Not X = m_CurrentImage Then
            Set m_Image(y) = mArray(X)
            y = y + 1
        End If
    Next X
    
handle:
    If m_CurrentImage > ImageCount - 1 Then m_CurrentImage = Me.ImageCount - 1
    RedrawControl
End Property

'Property to Move Images
Public Property Get MoveImageTo() As Long
    MoveImageTo = -1
End Property

Public Property Let MoveImageTo(ByVal vNewValue As Long)
Dim TmpImage As StdPicture

    If vNewValue > Me.ImageCount - 1 Then vNewValue = ImageCount - 1
    Set TmpImage = m_Image(m_CurrentImage)
    Set m_Image(m_CurrentImage) = m_Image(vNewValue)
    Set m_Image(vNewValue) = TmpImage
    RedrawControl
    
End Property

' CyclePic - For sequential use or animation purpose
Public Function CyclePic(Optional ByVal blnUp As Boolean = True) As StdPicture

    Select Case blnUp
        Case True
            If CurrentImage = ImageCount - 1 Then
                CurrentImage = 0
            Else
                CurrentImage = CurrentImage + 1
            End If
        Case False
            If CurrentImage = 0 Then
                CurrentImage = ImageCount - 1
            Else
                CurrentImage = CurrentImage - 1
            End If
    End Select

  Set CyclePic = ListImages(CurrentImage)

End Function

'Draw the control
Private Sub RedrawControl()
Dim mHeight As Long
Dim mWidth As Double
Dim Rct As RECT
Dim X As Long
    
    If Ambient.UserMode Then Exit Sub
    Cls
    m_ImageCount = Me.ImageCount
    If m_ImageCount = 0 Then Exit Sub
    Debug.Print "Redrawing..."
    
    mWidth = (ScaleWidth - 1) / m_ImageCount
    mHeight = ScaleHeight - TextHeight("A")
    Rct.Top = mHeight
    Rct.Bottom = ScaleHeight
    
    For X = 0 To m_ImageCount - 1
        Rct.Left = X * mWidth
        Rct.Right = (X + 1) * mWidth
        If X = m_CurrentImage Then Rectangle hDC, Rct.Left, 0, Rct.Right + 1, ScaleHeight
        PaintPicture m_Image(X), X * mWidth + 2, 2, mWidth - 2, mHeight - 2
        DrawText hDC, X, -1, Rct, 1
    Next X
       
End Sub

