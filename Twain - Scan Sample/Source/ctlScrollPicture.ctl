VERSION 5.00
Begin VB.UserControl ScrollPicture 
   ClientHeight    =   2700
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3180
   ClipControls    =   0   'False
   LockControls    =   -1  'True
   ScaleHeight     =   180
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   212
   ToolboxBitmap   =   "ctlScrollPicture.ctx":0000
   Begin VB.PictureBox PicScroll 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2355
      Left            =   0
      ScaleHeight     =   157
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   156
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   2340
   End
   Begin VB.PictureBox PicLoaded 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   285
      ScaleHeight     =   158
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   150
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   195
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Image cGrab 
      Height          =   480
      Left            =   2670
      Picture         =   "ctlScrollPicture.ctx":0312
      Top             =   150
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image cRelease 
      Height          =   480
      Left            =   2670
      Picture         =   "ctlScrollPicture.ctx":061C
      Top             =   630
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "ScrollPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Rem +++                                                                                                   +++
Rem +++ I didnt make this Control, i have modifyed a small bit of the Code...                             +++
Rem +++ Credit goes to :  Carles P.V. - in 2001 - Created Picture control.                                +++
Rem +++ Email HIM not me for Questions ect: carles_pv@terra.es                                                   +++
Rem +++                                                                                                   +++
Rem +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type Sizes
             sWidth As Single
             sHeight As Single
End Type

Private p_HV As Integer
Private p_VV As Integer
Private p_HM As Integer
Private p_VM As Integer

Private p_PictureMoving As Boolean
Private p_ancPos As POINTAPI
Private p_tmpPos As POINTAPI

Private p_ZoomFactor(0 To 14) As Integer
Private p_ZoomIndex As Integer

Private p_Size As Sizes
Private p_MouseIcon

Public Event Click()
Public Event DblClick()
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Private Const SRCCOPY        As Long = &HCC0020
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long


Private Sub UserControl_InitProperties()
On Error Resume Next

    BackColor = &H8000000F
    
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error Resume Next

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    
    PicScroll.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error Resume Next

    Call PropBag.WriteProperty("BackColor", PicScroll.BackColor, &H8000000F)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)

End Sub

Private Sub UserControl_Initialize()
On Error Resume Next

    ReadZoomFactors
    
End Sub

Public Sub About()
On Error Resume Next
    MsgBox "This Program contains implements of Brush Stroke, or" & vbNewLine & "One or More of the Brush Stroke Components." & vbNewLine & vbNewLine & "Copyright © 2005 DaTo Software® All Rights Reserved." & vbNewLine & "Other marks belong to their respective owners." & vbNewLine & vbNewLine & "Image Components Taken From Brush preview's / Filters."
End Sub

Private Sub UserControl_Show()
On Error Resume Next

    UserControl_Resize
    
End Sub

Private Sub UserControl_Resize()
On Error Resume Next

    With PicScroll
    
        p_HM = IIf(Picture <> 0, .ScaleWidth - ScaleWidth, 0)
        p_VM = IIf(Picture <> 0, .ScaleHeight - ScaleHeight, 0)
        
        If Picture <> 0 Then
            p_HV = p_HM / 2
            p_VV = p_VM / 2
            .Move -p_HV, -p_VV
        Else
            .Width = ScaleWidth
            .Height = ScaleHeight
        End If
        
        If .Width > ScaleWidth Or .Height > ScaleHeight Then
            PicScroll.MousePointer = vbCustom
            PicScroll.MouseIcon = cRelease
        Else
            PicScroll.MousePointer = vbDefault
        End If
        
        .Visible = True
        
    End With
    
End Sub

Private Sub UserControl_Terminate()
On Error Resume Next

    Set Picture = Nothing
    
End Sub

Public Property Get BackColor() As OLE_COLOR
On Error Resume Next

    BackColor = UserControl.BackColor
    
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
On Error Resume Next

    UserControl.BackColor = New_BackColor
    PicScroll.BackColor = New_BackColor
    PicLoaded.BackColor = New_BackColor
    PropertyChanged "BackColor"
    
End Property

Public Property Get Enabled() As Boolean
On Error Resume Next

    Enabled = UserControl.Enabled
    
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
On Error Resume Next

    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    
End Property

Public Property Get Picture() As StdPicture
On Error Resume Next

    Set Picture = PicScroll
    
End Property

Public Property Set Picture(ByVal New_Picture As StdPicture)
On Error Resume Next

    PicScroll.Visible = False
    
    Set PicScroll = New_Picture
    Set PicLoaded = New_Picture
    
    p_Size.sWidth = PicScroll.Width
    p_Size.sHeight = PicScroll.Height
    
    UserControl_Resize
    
    PropertyChanged "Picture"
    
End Property

Public Property Get ZoomPercent()
On Error Resume Next

    ZoomPercent = p_ZoomFactor(p_ZoomIndex)
    
End Property


Private Sub PicScroll_Click()
On Error Resume Next

    RaiseEvent Click
    
End Sub

Private Sub PicScroll_DblClick()
On Error Resume Next

    RaiseEvent DblClick
    
End Sub

Private Sub PicScroll_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    p_PictureMoving = True
    If Button = vbLeftButton Then PicScroll.MouseIcon = cGrab

    GetCursorPos p_ancPos
    p_tmpPos.x = p_ancPos.x
    p_tmpPos.y = p_ancPos.y
    
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub

Private Sub PicScroll_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Button <> vbLeftButton Or _
       Picture = 0 Or _
       p_PictureMoving = False Then Exit Sub
       
    GetCursorPos p_ancPos
    
    If PicScroll.ScaleWidth > ScaleWidth Then
        If (p_ancPos.x - p_tmpPos.x) > 0 Then
            If p_HV - (p_ancPos.x - p_tmpPos.x) > 0 Then
                p_HV = p_HV - (p_ancPos.x - p_tmpPos.x)
            Else
                p_HV = 0
            End If
        Else
            If p_HV - (p_ancPos.x - p_tmpPos.x) < p_HM Then
                p_HV = p_HV - (p_ancPos.x - p_tmpPos.x)
            Else
                p_HV = p_HM
            End If
        End If
    End If
    
    If PicScroll.ScaleHeight > ScaleHeight Then
        If (p_ancPos.y - p_tmpPos.y) > 0 Then
            If p_VV - (p_ancPos.y - p_tmpPos.y) > 0 Then
                p_VV = p_VV - (p_ancPos.y - p_tmpPos.y)
            Else
                p_VV = 0
            End If
        Else
            If p_VV - (p_ancPos.y - p_tmpPos.y) < p_VM Then
                p_VV = p_VV - (p_ancPos.y - p_tmpPos.y)
            Else
                p_VV = p_VM
            End If
        End If
    End If
    
    p_tmpPos.x = p_ancPos.x
    p_tmpPos.y = p_ancPos.y
    
    PicScroll.Move -p_HV, -p_VV
    
    RaiseEvent MouseMove(Button, Shift, x, y)
    
End Sub

Private Sub PicScroll_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    p_PictureMoving = False
    PicScroll.MouseIcon = cRelease
    
    RaiseEvent MouseUp(Button, Shift, x, y)
    
End Sub

Public Sub Clear()
On Error Resume Next

    Set Picture = Nothing
    
End Sub

Public Sub BestFit()
On Error Resume Next

    If Picture = 0 Then Exit Sub
    
    Dim relW As Single
    Dim relH As Single
    
    Dim W As Integer
    Dim h As Integer
    
    If (p_Size.sHeight <> ScaleHeight) Or (p_Size.sWidth <> ScaleWidth) Then
    
        relH = ScaleHeight / p_Size.sHeight
        relW = ScaleWidth / p_Size.sWidth
        
        If relW < relH Then
           W = ScaleWidth
           h = Int(p_Size.sHeight * relW)
        Else
           h = ScaleHeight
           W = Int(p_Size.sWidth * relH)
        End If
        
    Else
    
        W = p_Size.sWidth
        h = p_Size.sHeight
        
    End If
    
    PictureSize W, h

End Sub

Public Sub z_Stretch()
On Error Resume Next
    If Picture = 0 Then Exit Sub

    PictureSize ScaleWidth, ScaleHeight

End Sub

Public Sub z_ZoomIn()
On Error Resume Next

    If Picture = 0 Then Exit Sub
    
    If p_ZoomIndex < 14 Then
        p_ZoomIndex = p_ZoomIndex + 1
        PictureSize p_Size.sWidth * p_ZoomFactor(p_ZoomIndex) / 100, p_Size.sHeight * p_ZoomFactor(p_ZoomIndex) / 100
    End If
    
End Sub

Public Sub z_ZoomOut()
On Error Resume Next

    If Picture = 0 Then Exit Sub
    
    If p_ZoomIndex > 0 Then
        p_ZoomIndex = p_ZoomIndex - 1
        PictureSize p_Size.sWidth * p_ZoomFactor(p_ZoomIndex) / 100, p_Size.sHeight * p_ZoomFactor(p_ZoomIndex) / 100
    End If
    
End Sub

Public Sub z_ZoomPrevious()
On Error Resume Next

    If Picture = 0 Then Exit Sub
    
    PictureSize p_Size.sWidth * p_ZoomFactor(p_ZoomIndex) / 100, p_Size.sHeight * p_ZoomFactor(p_ZoomIndex) / 100
    
End Sub

Public Sub z_ZoomActualPixels()
On Error Resume Next

    If Picture = 0 Then Exit Sub
    p_ZoomIndex = 10
    PictureSize p_Size.sWidth, p_Size.sHeight
    
End Sub

Private Sub PictureSize(ByVal NewWidth As Integer, ByVal NewHeight As Integer)
    
    Screen.MousePointer = vbHourglass
    PicScroll.Visible = False
    
        PicScroll.Width = NewWidth
        PicScroll.Height = NewHeight
       
        On Error Resume Next
        StretchBlt PicScroll.hdc, _
                   0, 0, _
                   NewWidth, NewHeight, _
                   PicLoaded.hdc, _
                   0, 0, _
                   p_Size.sWidth, p_Size.sHeight, _
                   SRCCOPY
                       
        If Err.Number > 0 Then
            Err.Clear
            z_ZoomActualPixels
        End If
    
    UserControl_Resize
    Screen.MousePointer = vbDefault

End Sub

Private Sub ReadZoomFactors()
On Error Resume Next
    p_ZoomFactor(0) = 5
    p_ZoomFactor(1) = 10
    p_ZoomFactor(2) = 20
    p_ZoomFactor(3) = 30
    p_ZoomFactor(4) = 40
    p_ZoomFactor(5) = 50
    p_ZoomFactor(6) = 60
    p_ZoomFactor(7) = 70
    p_ZoomFactor(8) = 80
    p_ZoomFactor(9) = 90
    p_ZoomFactor(10) = 100
    p_ZoomFactor(11) = 125
    p_ZoomFactor(12) = 150
    p_ZoomFactor(13) = 175
    p_ZoomFactor(14) = 200
    
    p_ZoomIndex = 10
    
End Sub
