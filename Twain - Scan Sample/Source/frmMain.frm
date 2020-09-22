VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TWAIN Sample in VB."
   ClientHeight    =   6255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog dlgSave 
      Left            =   7080
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TWAIN.ScrollPicture ScrollPicture1 
      Height          =   6015
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   10610
      BackColor       =   16777215
      BackColor       =   16777215
      Picture         =   "frmMain.frx":000C
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7080
      TabIndex        =   3
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Picture"
      Height          =   375
      Left            =   7080
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton cmdAcquire 
      Caption         =   "TWAIN Acquire"
      Height          =   375
      Left            =   7080
      TabIndex        =   1
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "TWAIN Select"
      Height          =   375
      Left            =   7080
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Line Line1 
      X1              =   6960
      X2              =   6960
      Y1              =   120
      Y2              =   6120
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Rem +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Rem +++ In the past few Years on PSC there have been a few submissions about Scanning...                  +++
Rem +++ I Have Looked and Looked for a decent submission about Scanning                                   +++
Rem +++ and All the Functions in the TWAIN32 Dlls.                                                        +++
Rem +++                                                                                                   +++
Rem +++ The other day i came across a website that Coded a TWAIN32DLL in C++                              +++
Rem +++ All that this app does is 'Wrap Arround' the C++ Function Calls to VB                             +++
Rem +++                                                                                                   +++
Rem +++ There are Few Example for VB programmers on how to TWAIN ect... So please Learn from This One.    +++
Rem +++ The origonal C++ Source can be Downloaded from www.dosadi.com                                     +++
Rem +++                                                                                                   +++
Rem +++ The User Control i didnt Create, please refer to that for more info.                              +++
Rem +++                                                                                                   +++
Rem +++ Please Leave your Comments ect... on PS Code, where you downloaded this From.                     +++
Rem +++ David Nedved.                                                                                     +++
Rem +++                                                                                                   +++
Rem +++ I Have tryed to Comment each Module as Best that i can, but i am verry 'Slack' when it comes      +++
Rem +++ To Comenting... So Excuse my Language ect...                                                      +++
Rem +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Private Sub cmdAcquire_Click()
'Dim what we will need to Scan The Pic...
Dim Ret As Long, PictureFile As String
'Select where we want to Save the Temp Picture File To.
PictureFile = App.Path & "\~" & App.hInstance & ".tmp"
'PicturFile is the temporary file i just use the apps HInstance for Temp files,
'when im working with more than one app... :)
'In "temp.bmp" the image will stored until the end of the action
Ret = TWAIN_AcquireToFilename(Me.HWND, PictureFile)
If Ret = 0 Then
 Set Me.ScrollPicture1.Picture = LoadPicture(PictureFile)
 'Load the temporary picture file
 Kill PictureFile
 'Delete the temporary picture file
Else
 MsgBox "Scan unsuccessful" & vbNewLine & "Please check that your Scanner is switched On, and that you have Selected the right Scanner.", vbExclamation, "TWAIN Acquire Failed."
End If
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdSave_Click()
'Save the Picture File
'I could use Calls for the Dialoug Box, but i dont want to make this to Confusing...
'So i will use a Microsoft Common Dialog Control (Comes Default with VB 5, VB 6)
'I Always Use a With Statement When Using a MS Dialog beacuse it makes it easier
'To Understand, and its not all crammed in...
'You should be able to figer out what this does... Its Not that hard :)
'Just in case...
On Error Resume Next
Dim sFileName As String
With Me.dlgSave
 .DialogTitle = "Save Image File"
 .Filter = "Bitmap Files (*.bmp)|*.bmp|Gif Files (*.gif)|*.gif"
 .ShowSave
 sFileName = .FileName
End With
SavePicture Me.ScrollPicture1.Picture, sFileName
End Sub

Private Sub cmdSelect_Click()
TWAIN_SelectImageSource (Me.HWND)
End Sub

