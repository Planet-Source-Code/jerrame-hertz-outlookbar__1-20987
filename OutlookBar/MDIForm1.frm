VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Outlook Shortcut Bar"
   ClientHeight    =   7590
   ClientLeft      =   6315
   ClientTop       =   4620
   ClientWidth     =   8925
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   3  'Align Left
      BackColor       =   &H8000000C&
      Height          =   7590
      Left            =   0
      ScaleHeight     =   7530
      ScaleWidth      =   1695
      TabIndex        =   0
      Top             =   0
      Width           =   1760
      Begin VB.Frame Frame1 
         BackColor       =   &H8000000C&
         BorderStyle     =   0  'None
         Height          =   6735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   1695
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000C&
            BorderStyle     =   0  'None
            Height          =   6375
            Left            =   0
            TabIndex        =   3
            Top             =   360
            Width           =   1695
            Begin VB.CommandButton Command4 
               BackColor       =   &H8000000C&
               Height          =   735
               Index           =   3
               Left            =   360
               Picture         =   "MDIForm1.frx":0000
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   4800
               Width           =   855
            End
            Begin VB.CommandButton Command4 
               BackColor       =   &H8000000C&
               Height          =   735
               Index           =   2
               Left            =   360
               Picture         =   "MDIForm1.frx":030A
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   3360
               Width           =   855
            End
            Begin VB.CommandButton Command4 
               BackColor       =   &H8000000C&
               Height          =   735
               Index           =   1
               Left            =   360
               Picture         =   "MDIForm1.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   14
               Top             =   1920
               Width           =   855
            End
            Begin VB.CommandButton Command4 
               BackColor       =   &H8000000C&
               Height          =   735
               Index           =   0
               Left            =   360
               Picture         =   "MDIForm1.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   13
               Top             =   600
               Width           =   855
            End
            Begin VB.CommandButton Command2 
               Caption         =   "Tab 2"
               Height          =   375
               Left            =   0
               TabIndex        =   4
               Top             =   0
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Button 4"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   20
               Top             =   5640
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Button 3"
               Height          =   255
               Index           =   2
               Left            =   240
               TabIndex        =   19
               Top             =   4200
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Button 2"
               Height          =   375
               Index           =   1
               Left            =   240
               TabIndex        =   18
               Top             =   2760
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Button 1"
               Height          =   255
               Index           =   0
               Left            =   240
               TabIndex        =   17
               Top             =   1440
               Width           =   1095
            End
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000C&
            Height          =   735
            Index           =   3
            Left            =   360
            Picture         =   "MDIForm1.frx":0C28
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   4680
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000C&
            Height          =   735
            Index           =   2
            Left            =   360
            Picture         =   "MDIForm1.frx":0F32
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3360
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000C&
            Height          =   735
            Index           =   1
            Left            =   360
            Picture         =   "MDIForm1.frx":123C
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2040
            Width           =   855
         End
         Begin VB.CommandButton Command3 
            BackColor       =   &H8000000C&
            Height          =   735
            Index           =   0
            Left            =   360
            Picture         =   "MDIForm1.frx":1546
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   720
            Width           =   855
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Tab 1"
            Height          =   375
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Button 4"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   12
            Top             =   5520
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Button 3"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   11
            Top             =   4200
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Button 2"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   10
            Top             =   2880
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Button 1"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   1560
            Width           =   1095
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuFileBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save &As..."
      End
      Begin VB.Menu mnuFileSaveAll 
         Caption         =   "Save A&ll"
      End
      Begin VB.Menu mnuFileBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "Propert&ies"
      End
      Begin VB.Menu mnuFileBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintSetup 
         Caption         =   "Print Set&up..."
      End
      Begin VB.Menu mnuFilePrintPreview 
         Caption         =   "Print Pre&view"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "&Print..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileBar4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSend 
         Caption         =   "Sen&d..."
      End
      Begin VB.Menu mnuFileBar5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "&Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEditPasteSpecial 
         Caption         =   "Paste &Special..."
      End
      Begin VB.Menu mnuEditBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDSelectAll 
         Caption         =   "Select &All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditInvertSelection 
         Caption         =   "&Invert Selection"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "&Toolbar"
      End
      Begin VB.Menu mnuViewStatusBar 
         Caption         =   "Status &Bar"
      End
      Begin VB.Menu mnuViewBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewLargeIcons 
         Caption         =   "Lar&ge Icons"
      End
      Begin VB.Menu mnuViewSmallIcons 
         Caption         =   "S&mall Icons"
      End
      Begin VB.Menu mnuViewList 
         Caption         =   "&List"
      End
      Begin VB.Menu mnuViewDetails 
         Caption         =   "&Details"
      End
      Begin VB.Menu mnuViewBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewArrangeIcons 
         Caption         =   "Arrange &Icons"
         Begin VB.Menu mnuVAIByName 
            Caption         =   "by &Name"
         End
         Begin VB.Menu mnuVAIByType 
            Caption         =   "by &Type"
         End
         Begin VB.Menu mnuVAIBySize 
            Caption         =   "by Si&ze"
         End
         Begin VB.Menu mnuVAIByDate 
            Caption         =   "by &Date"
         End
      End
      Begin VB.Menu mnuViewLineUpIcons 
         Caption         =   "Li&ne Up Icons"
      End
      Begin VB.Menu mnuViewBar3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "&Refresh"
      End
      Begin VB.Menu mnuViewOptions 
         Caption         =   "&Options..."
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowNewWindow 
         Caption         =   "&New Window"
      End
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
      Begin VB.Menu mnuWindowBar1 
         Caption         =   "-"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub Command1_Click()
    MoveBar
End Sub


Private Sub Command2_Click()
    MoveBar
End Sub

Private Sub MoveBar()
    Select Case Frame2.Top
        Case 360
            Frame2.Move 0, 5880, 1695, 375
        Case 5880
            Frame2.Move 0, 360, 1695, 6255
    End Select
End Sub


