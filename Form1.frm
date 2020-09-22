VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "XP Button"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4710
   ClipControls    =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   215
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   314
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboStyle 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2700
      Width           =   1545
   End
   Begin VB.OptionButton Option1 
      Caption         =   "OptionButton"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1980
      Value           =   -1  'True
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      Caption         =   "CheckBox"
      Enabled         =   0   'False
      Height          =   285
      Index           =   5
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1980
      Value           =   2  'Grayed
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2430
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "OptionButton"
      Height          =   285
      Index           =   4
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1620
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "OptionButton"
      Height          =   285
      Index           =   3
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1260
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "OptionButton"
      Height          =   285
      Index           =   2
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   900
      Width           =   1365
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "OptionButton"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   540
      Width           =   1365
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CheckBox"
      Height          =   285
      Index           =   4
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1620
      Value           =   2  'Grayed
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Checkbox"
      Height          =   285
      Index           =   2
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   900
      Value           =   2  'Grayed
      Width           =   1185
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Checkbox"
      Height          =   285
      Index           =   0
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   180
      Value           =   2  'Grayed
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1980
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1530
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   180
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "CheckBox"
      Height          =   285
      Index           =   3
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1260
      Value           =   2  'Grayed
      Width           =   1185
   End
   Begin VB.OptionButton Option1 
      Caption         =   "OptionButton"
      Height          =   285
      Index           =   0
      Left            =   1530
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   180
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3060
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   630
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Checkbox"
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   1
      Left            =   180
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Value           =   2  'Grayed
      Width           =   1185
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboStyle_Click()
   IDB_CHECKBOX = 102
   Select Case cboStyle.ListIndex
   Case 0: IDB_BUTTON = 101
   Case 1: IDB_BUTTON = 105
   Case 2: IDB_BUTTON = 106
   Case 3
      IDB_BUTTON = 107
      IDB_CHECKBOX = 104
   Case 4: IDB_BUTTON = 108
   Case 5: IDB_BUTTON = 109
   End Select
   Me.Refresh
End Sub

Private Sub Form_Load()
   Dim i As Integer
   
   cboStyle.AddItem "Blue"
   cboStyle.AddItem "Olive Green"
   cboStyle.AddItem "Silver"
   cboStyle.AddItem "Tasblue"
   cboStyle.AddItem "Gold"
   cboStyle.AddItem "Aqua"
   cboStyle.ListIndex = 0
   HookForm hWnd
   For i = 0 To 4
      HookControl Command1(i).hWnd
      HookControl Check1(i).hWnd
      HookControl Option1(i).hWnd
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim i As Integer
   
   UnHook hWnd
   For i = 0 To 4
      UnHook Command1(i).hWnd
      UnHook Check1(i).hWnd
      UnHook Option1(i).hWnd
   Next
End Sub

Public Function GetControl(hWnd As Long) As Control
   Dim hWndItem As Long
   Dim Ctrl As Control
  
   On Error Resume Next
   For Each Ctrl In Form1.Controls
      hWndItem = Ctrl.hWnd
      If Err.Number = 0 Then
         If hWndItem = hWnd Then
            Set GetControl = Ctrl
            Exit Function
         End If
      End If
   Next
End Function
