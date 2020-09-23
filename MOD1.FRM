VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form frm1 
   Caption         =   "MultiDoc     Blank.txt"
   ClientHeight    =   5805
   ClientLeft      =   705
   ClientTop       =   990
   ClientWidth     =   8280
   Icon            =   "Mod1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5805
   ScaleWidth      =   8280
   WindowState     =   2  'Maximized
   Begin VB.Frame fraNew 
      Caption         =   "Create New File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   930
      Left            =   1965
      TabIndex        =   27
      Top             =   2220
      Visible         =   0   'False
      Width           =   4605
      Begin VB.CommandButton cmdOK 
         Caption         =   "&Cancel"
         Height          =   315
         Index           =   1
         Left            =   3270
         TabIndex        =   31
         Top             =   525
         Width           =   1230
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   315
         Index           =   0
         Left            =   3270
         TabIndex        =   30
         Top             =   195
         Width           =   1230
      End
      Begin VB.ComboBox cboNew 
         Height          =   315
         ItemData        =   "Mod1.frx":058A
         Left            =   1590
         List            =   "Mod1.frx":0594
         TabIndex        =   29
         Text            =   "TXT"
         ToolTipText     =   "Extension"
         Top             =   195
         Width           =   1500
      End
      Begin VB.TextBox txtNew 
         Height          =   315
         Left            =   60
         TabIndex        =   28
         Text            =   "Blank"
         ToolTipText     =   "FileName"
         Top             =   195
         Width           =   1500
      End
   End
   Begin VB.Frame fraTools 
      Caption         =   "Tools"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   1725
      Begin VB.CommandButton cmdClear 
         Height          =   315
         Left            =   1320
         Picture         =   "Mod1.frx":05A2
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "New"
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cmdCase 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1320
         Picture         =   "Mod1.frx":08E4
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "lowercase"
         Top             =   200
         Width           =   315
      End
      Begin VB.CommandButton cmdCase 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1005
         Picture         =   "Mod1.frx":0C26
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "UPPERCASE"
         Top             =   200
         Width           =   315
      End
      Begin VB.OptionButton optAlign 
         Height          =   315
         Index           =   0
         Left            =   60
         Picture         =   "Mod1.frx":0F68
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Left Align"
         Top             =   200
         Value           =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton optAlign 
         Height          =   315
         Index           =   2
         Left            =   375
         Picture         =   "Mod1.frx":12AA
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Center"
         Top             =   200
         Width           =   315
      End
      Begin VB.OptionButton optAlign 
         Height          =   315
         Index           =   1
         Left            =   690
         Picture         =   "Mod1.frx":15EC
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Right Align"
         Top             =   200
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   0
         Left            =   1320
         Picture         =   "Mod1.frx":192E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Bullet Indent"
         Top             =   510
         Width           =   315
      End
      Begin VB.VScrollBar vsbSize 
         Height          =   315
         LargeChange     =   4
         Left            =   1005
         Max             =   2
         Min             =   100
         TabIndex        =   26
         Top             =   825
         Value           =   10
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   5
         Left            =   1005
         Picture         =   "Mod1.frx":1C70
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "StrikeThru"
         Top             =   510
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   4
         Left            =   375
         Picture         =   "Mod1.frx":1FB2
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Color"
         Top             =   825
         Width           =   315
      End
      Begin VB.CommandButton cmdFont 
         Height          =   315
         Left            =   60
         Picture         =   "Mod1.frx":22F4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Font"
         Top             =   825
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   3
         Left            =   690
         Picture         =   "Mod1.frx":2636
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Underline"
         Top             =   510
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   2
         Left            =   375
         Picture         =   "Mod1.frx":2978
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Italic"
         Top             =   510
         Width           =   315
      End
      Begin VB.CheckBox chk 
         Height          =   315
         Index           =   1
         Left            =   60
         Picture         =   "Mod1.frx":2CBA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Boldface"
         Top             =   510
         Width           =   315
      End
      Begin VB.Label lblFont 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   60
         TabIndex        =   32
         ToolTipText     =   "Current Font Name"
         Top             =   1140
         Width           =   1590
      End
      Begin VB.Label lblSize 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   690
         TabIndex        =   25
         ToolTipText     =   "Text Size"
         Top             =   825
         Width           =   315
      End
   End
   Begin TabDlg.SSTab tab1 
      Height          =   3960
      Left            =   0
      TabIndex        =   11
      Top             =   1500
      Width           =   7440
      _ExtentX        =   13123
      _ExtentY        =   6985
      _Version        =   327680
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Blank.TXT"
      TabPicture(0)   =   "Mod1.frx":2FFC
      Tab(0).ControlCount=   1
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtb(0)"
      Tab(0).Control(0).Enabled=   0   'False
      TabCaption(1)   =   "Blank.TXT"
      TabPicture(1)   =   "Mod1.frx":3018
      Tab(1).ControlCount=   1
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "rtb(1)"
      Tab(1).Control(0).Enabled=   0   'False
      TabCaption(2)   =   "Blank.TXT"
      TabPicture(2)   =   "Mod1.frx":3034
      Tab(2).ControlCount=   1
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "rtb(2)"
      Tab(2).Control(0).Enabled=   0   'False
      TabCaption(3)   =   "Blank.TXT"
      TabPicture(3)   =   "Mod1.frx":3050
      Tab(3).ControlCount=   1
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "rtb(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2340
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   375
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4128
         _Version        =   327680
         HideSelection   =   0   'False
         ScrollBars      =   3
         RightMargin     =   4989
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Mod1.frx":306C
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2340
         Index           =   1
         Left            =   -74985
         TabIndex        =   12
         Top             =   315
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4128
         _Version        =   327680
         HideSelection   =   0   'False
         ScrollBars      =   3
         RightMargin     =   4989.764
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Mod1.frx":314E
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2340
         Index           =   2
         Left            =   -74970
         TabIndex        =   13
         Top             =   345
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4128
         _Version        =   327680
         HideSelection   =   0   'False
         ScrollBars      =   3
         RightMargin     =   4989.764
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Mod1.frx":3230
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtb 
         Height          =   2340
         Index           =   3
         Left            =   -75000
         TabIndex        =   14
         Top             =   330
         Width           =   4530
         _ExtentX        =   7990
         _ExtentY        =   4128
         _Version        =   327680
         HideSelection   =   0   'False
         ScrollBars      =   3
         RightMargin     =   4989.764
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"Mod1.frx":3312
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame fraSearch 
      Caption         =   "Search and Replace"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1500
      Left            =   1740
      TabIndex        =   3
      Top             =   0
      Width           =   6450
      Begin VB.Frame fraS 
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   540
         Left            =   60
         TabIndex        =   15
         Top             =   915
         Width           =   3210
         Begin VB.OptionButton opt 
            Caption         =   "All Text"
            Height          =   285
            Index           =   1
            Left            =   1845
            TabIndex        =   17
            Top             =   195
            Width           =   1035
         End
         Begin VB.OptionButton opt 
            Caption         =   "Current Text Box"
            Height          =   285
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   195
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "Find Again"
         Height          =   315
         Left            =   3915
         TabIndex        =   8
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton cmdR 
         Caption         =   "Replace"
         Enabled         =   0   'False
         Height          =   315
         Left            =   3915
         TabIndex        =   7
         Top             =   600
         Width           =   1230
      End
      Begin VB.CommandButton cmdRAll 
         Caption         =   "Replace All"
         Enabled         =   0   'False
         Height          =   315
         Left            =   5160
         TabIndex        =   6
         Top             =   600
         Width           =   1230
      End
      Begin VB.CheckBox chk1 
         Caption         =   "Case sensitive"
         Height          =   240
         Left            =   3915
         TabIndex        =   5
         Top             =   975
         Width           =   2040
      End
      Begin VB.CheckBox chk2 
         Caption         =   "Whole word only"
         Height          =   240
         Left            =   3915
         TabIndex        =   4
         Top             =   1215
         Width           =   2040
      End
      Begin VB.TextBox txtF 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         TabIndex        =   1
         Top             =   225
         Width           =   2115
      End
      Begin VB.TextBox txtR 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1725
         TabIndex        =   2
         Top             =   600
         Width           =   2115
      End
      Begin VB.Label lbl1 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Find what"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Width           =   3825
      End
      Begin VB.Label lbl2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Replace with"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   60
         TabIndex        =   9
         Top             =   555
         Width           =   3825
      End
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   105
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
      FontSize        =   1.17491e-38
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuReOpen 
         Caption         =   "&ReOpen"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuRename 
         Caption         =   "ReNa&me"
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save &As"
      End
      Begin VB.Menu mnu4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPrintSelected 
         Caption         =   "Print Selected &Text"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "&Print Current Text Box"
      End
      Begin VB.Menu mnu5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuClear 
         Caption         =   "Clea&r"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear &All"
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Co&py"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCut 
         Caption         =   "Cu&t"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "&Paste"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSelect 
         Caption         =   "&Select All"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const Filt = "Text  (*.TXT)|*.TXT|Rich Text  (*.RTF)|*.RTF|All Files  (*.*)|*.*"
Dim Pos(3) As Integer
Dim OpenFile(3) As String, SavedAs(3) As Boolean
Dim FindFlags As Integer
Dim FName As String, ReplaceYes As Boolean
Private Sub chk_Click(Index As Integer)
On Error Resume Next
Select Case Index
Case 0  'bullet
  rtb(tab1.Tab).SelBullet = chk(0).Value
Case 1  'Bold
  rtb(tab1.Tab).SelBold = chk(1).Value
Case 2  'Italic
  rtb(tab1.Tab).SelItalic = chk(2).Value
Case 3  'Underline
  rtb(tab1.Tab).SelUnderline = chk(3).Value
Case 4  'Color
  If chk(4).Value = 1 Then
    cd1.ShowColor
    rtb(tab1.Tab).SelColor = cd1.Color
  Else
    cd1.Color = vbBlack
    rtb(tab1.Tab).SelColor = cd1.Color
  End If
Case 5  'StrikeThru
  rtb(tab1.Tab).SelStrikeThru = chk(5).Value
End Select
End Sub
Private Sub cmdCase_Click(Index As Integer)
If Index = 0 Then      'UPPERCASE
  rtb(tab1.Tab).SelText = UCase(rtb(tab1.Tab).SelText)
ElseIf Index = 1 Then  'lowercase
  rtb(tab1.Tab).SelText = LCase(rtb(tab1.Tab).SelText)
End If
rtb(tab1.Tab).SetFocus
End Sub
Private Sub cmdClear_Click()
mnuNew_Click
End Sub
Private Sub cmdFont_Click()
On Error Resume Next
'Finds 7 Properties
cd1.FontItalic = rtb(tab1.Tab).SelItalic
cd1.FontBold = rtb(tab1.Tab).SelBold
cd1.FontName = FName
cd1.FontSize = rtb(tab1.Tab).SelFontSize
cd1.FontStrikethru = rtb(tab1.Tab).SelStrikeThru
cd1.FontUnderline = rtb(tab1.Tab).SelUnderline
cd1.Color = rtb(tab1.Tab).SelColor
'Sets Flags and Shows FontSelect
cd1.Flags = cdlCFBoth Or cdlCFEffects Or 262144
cd1.ShowFont
'Returns 7 Properties
AdjustButtons
FName = cd1.FontName
lblFont.Caption = FName
vsbSize.Value = cd1.FontSize
rtb(tab1.Tab).SelColor = cd1.Color
End Sub
Private Sub cmdOK_Click(Index As Integer)
If Index = 0 Then 'OK
  Select Case fraNew.Caption
  Case "Create New File"
    tab1.Caption = txtNew.Text & "." & cboNew.Text
    OpenFile(tab1.Tab) = tab1.Caption
    rtb(tab1.Tab).Text = ""
    SavedAs(tab1.Tab) = False
  Case "Rename " To "Rename z"
    tab1.Caption = txtNew.Text & "." & cboNew.Text
    OpenFile(tab1.Tab) = tab1.Caption
  End Select
  Me.Caption = "MultiDoc     " & OpenFile(tab1.Tab)
End If
fraNew.Visible = False
End Sub
Private Sub mnuCopy_Click()
Clipboard.Clear
Clipboard.SetText rtb(tab1.Tab).SelText
End Sub
Private Sub mnuCut_Click()
Clipboard.Clear
Clipboard.SetText rtb(tab1.Tab).SelText
rtb(tab1.Tab).SelText = ""
End Sub
Private Sub mnuFile_Click()
mnuReOpen.Enabled = SavedAs(tab1.Tab)
mnuReOpen.Caption = "&Re-Open " & tab1.Caption
mnuSave.Enabled = SavedAs(tab1.Tab)
mnuSave.Caption = "&Save " & tab1.Caption
mnuRename.Caption = "Rena&me " & tab1.Caption
End Sub
Private Sub mnuNew_Click()
fraNew.Visible = True
End Sub
Private Sub mnuPaste_Click()
rtb(tab1.Tab).SelText = Clipboard.GetText()
End Sub
Private Sub mnuPrintSelected_Click()
Printer.Print rtb(tab1.Tab).SelText
Printer.EndDoc
End Sub
Private Sub mnuRename_Click()
fraNew.Caption = "Rename " & tab1.Caption
fraNew.Visible = True
txtNew.ToolTipText = "FileName"
cboNew.ToolTipText = "Extension"
End Sub
Private Sub mnuReOpen_Click()
On Error GoTo FileError:
ResetButtons
'Adjusts For File Types
If UCase(Right(OpenFile(tab1.Tab), 3)) = "RTF" Then
  tmode = rtfRTF
Else
  tmode = rtfText
End If
rtb(tab1.Tab).LoadFile OpenFile(tab1.Tab), tmode
SavedAs(tab1.Tab) = True
Exit Sub
FileError:
If Err.Number = cdlCancel Then Exit Sub
MsgBox "Unkown error while opening file " & cd1.filename
OpenFile(tab1.Tab) = "Blank.TXT"
End Sub
Private Sub mnuSave_Click()
Dim Extension As String
Extension = UCase(Right(tab1.Caption, 3))
'This Adjusts Automatically for Saving Files
If Extension = "RTF" Then
  rtb(tab1.Tab).SaveFile OpenFile(tab1.Tab), rtfRTF
Else
  rtb(tab1.Tab).SaveFile OpenFile(tab1.Tab), rtfText
End If
End Sub
Private Sub mnuSelect_Click()
rtb(tab1.Tab).SelStart = 0
rtb(tab1.Tab).SelLength = Len(rtb(tab1.Tab).Text)
End Sub
Private Sub mnuExit_Click()
End
End Sub
Private Sub mnuOpen_Click()
On Error GoTo FileError:
cd1.CancelError = True
cd1.Flags = cdlOFNFileMustExist
cd1.Filter = Filt
cd1.ShowOpen
ResetButtons
'Adjusts For File Types
If UCase(Right(cd1.filename, 3)) = "RTF" Then
  tmode = rtfRTF
Else
  tmode = rtfText
End If
rtb(tab1.Tab).LoadFile cd1.filename, tmode
OpenFile(tab1.Tab) = cd1.filename
tab1.Caption = cd1.FileTitle
SavedAs(tab1.Tab) = True
Me.Caption = "MultiDoc     " & OpenFile(tab1.Tab)
Exit Sub
FileError:
If Err.Number = cdlCancel Then Exit Sub
MsgBox "Unkown error while opening file " & cd1.filename
OpenFile(tab1.Tab) = "Blank.TXT"
End Sub
Private Sub mnuSaveAs_Click()
Dim Extension As String
Extension = UCase(Right(tab1.Caption, 3))
On Error GoTo FileError:
'Auto Sets Extension
Select Case Extension
Case "TXT"
  cd1.FilterIndex = 1
Case "RTF"
  cd1.FilterIndex = 2
Case Else
  cd1.FilterIndex = 3
End Select
cd1.CancelError = True
cd1.filename = OpenFile(tab1.Tab)
cd1.Filter = Filt
cd1.ShowSave
'This Adjusts Automatically for Saving Files
If Extension = "RTF" Then
  rtb(tab1.Tab).SaveFile cd1.filename, rtfRTF
Else
  rtb(tab1.Tab).SaveFile cd1.filename, rtfText
End If
OpenFile(tab1.Tab) = cd1.filename
tab1.Caption = cd1.FileTitle
SavedAs(tab1.Tab) = True
Me.Caption = "MultiDoc     " & OpenFile(tab1.Tab)
Exit Sub
FileError:
If Err.Number = cdlCancel Then Exit Sub
MsgBox "Unkown error while opening file " & cd1.filename
OpenFile(tab1.Tab) = "Blank.TXT"
End Sub
Private Sub Form_Load()
Dim i As Integer
SizeMe
For i = 0 To 3
  Pos(i) = -1
  OpenFile(i) = "Blank.TXT"
Next
FName = txtF.FontName
lblFont.Caption = FName
If Command$ <> "" Then
  'Adjusts For File Types
  If UCase(Right(Command$, 3)) = "RTF" Then
    tmode = rtfRTF
  Else
    tmode = rtfText
  End If
  rtb(tab1.Tab).LoadFile Command$, tmode
  OpenFile(tab1.Tab) = Command$
  tab1.Caption = ExtractText(Command$, "\", False, False)
  SavedAs(tab1.Tab) = True
  Me.Caption = "MultiDoc     " & OpenFile(tab1.Tab)
End If
End Sub
Private Sub Form_Resize()
SizeMe
End Sub
Private Sub cmdNext_Click()
On Error Resume Next
SearchMe
'Found Something
If Pos(tab1.Tab) >= 0 Then
  cmdR.Enabled = True
  cmdRAll.Enabled = True
'Didn't Find
Else
  MsgBox "String not found"
  cmdR.Enabled = False
  cmdRAll.Enabled = False
End If
End Sub
Private Sub mnuAbout_Click()
MsgBox App.EXEName & " by Todd Spangler"
End Sub
Private Sub mnuClear_Click()
rtb(tab1.Tab).Text = ""
End Sub
Private Sub mnuClearAll_Click()
Dim i As Integer
For i = 0 To 3
  rtb(i).Text = ""
Next
End Sub
Private Sub mnuDelete_Click()
rtb(tab1.Tab).SelText = ""
End Sub
Private Sub mnuPrint_Click()
Printer.Print rtb(tab1.Tab).Text
Printer.EndDoc
End Sub
Private Sub optAlign_Click(Index As Integer)
rtb(tab1.Tab).SelAlignment = Index
End Sub
Private Sub cmdR_Click()
On Error Resume Next
'Searches & Replaces
rtb(tab1.Tab).SelText = txtR.Text
SearchMe
'Found Something
If Pos(tab1.Tab) >= 0 Then
  'Didn't Find
Else
  MsgBox "String not found"
  cmdR.Enabled = False
  cmdRAll.Enabled = False
End If
End Sub
Private Sub cmdRAll_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
BetterFind (txtF.Text)
If ReplaceYes Then
  ReplaceAll
End If
End Sub
Public Sub SizeMe()
Dim i As Integer
On Error Resume Next
For i = 0 To 3
  rtb(i).RightMargin = rtb(i).Width - 100
  rtb(i).Left = 60
  rtb(i).Top = 375
  tab1.Height = Me.ScaleHeight - fraSearch.Height
  tab1.Width = Me.ScaleWidth
  rtb(i).Width = tab1.Width - 120
  rtb(i).Height = tab1.Height - 450
  rtb(i).RightMargin = rtb(i).Width - 100
Next
End Sub
Private Sub rtb_Change(Index As Integer)
'Bullet Indent 0
rtb(tab1.Tab).SelBullet = chk(0).Value
'Bold 1
rtb(tab1.Tab).SelBold = chk(1).Value
'Italic 2
rtb(tab1.Tab).SelItalic = chk(2).Value
'Underline 3
rtb(tab1.Tab).SelUnderline = chk(3).Value
'Color 4
rtb(tab1.Tab).SelColor = cd1.Color
'StrikeThru
rtb(tab1.Tab).SelStrikeThru = chk(5).Value
'Size
rtb(tab1.Tab).SelFontSize = vsbSize.Value
'Font Name
rtb(tab1.Tab).SelFontName = FName
End Sub
Private Sub rtb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
'Select word, or sentence
If Shift = vbCtrlMask Then
  Select Case KeyCode
  'If Ctrl+S:
  Case vbKeyS
    rtb(tab1.Tab).Span ".?!", False, True
    SelectionStart = rtb(tab1.Tab).SelStart
    'Select to the end of the sentence.
    rtb(tab1.Tab).Span ".?!", True, True
    'Extend selection to include punctuation.
    SelectionEnd = rtb(tab1.Tab).SelStart + rtb(tab1.Tab).SelLength
    rtb(tab1.Tab).SelStart = SelectionStart
    rtb(tab1.Tab).SelLength = SelectionEnd - SelectionStart
    'If Ctrl+W:
  Case vbKeyW
    'Select to the end of the word.
    rtb(tab1.Tab).Span " ,;:.?!", False, True
    SelectionStart = rtb(tab1.Tab).SelStart
    'Select to the end of the word
    rtb(tab1.Tab).Span " ,;:.?!", True, True
    SelectionEnd = rtb(tab1.Tab).SelStart + rtb(tab1.Tab).SelLength
    rtb(tab1.Tab).SelStart = SelectionStart
    rtb(tab1.Tab).SelLength = SelectionEnd - SelectionStart
  End Select
End If
'Move pointer by word or sentence
If Shift = (vbCtrlMask Or vbShiftMask) Then
  Select Case KeyCode
  Case vbKeyS
    'Move pointer to end of sentence.
    rtb(tab1.Tab).UpTo ".?!", True, False
  Case vbKeyW
    'Move pointer to end of word.
    rtb(tab1.Tab).UpTo " ,;:.?!", True, False
  End Select
End If
End Sub
Public Sub SearchMe()
FindFlags = chk1.Value * 4 + chk2.Value * 2
On Error Resume Next
'Current Text Box
If opt(0).Value Then
  Pos(tab1.Tab) = rtb(tab1.Tab).Find(txtF.Text, Pos(tab1.Tab) + 1, , FindFlags)
  'All Text
ElseIf opt(1).Value Then
  Pos(tab1.Tab) = rtb(tab1.Tab).Find(txtF.Text, Pos(tab1.Tab) + 1, , FindFlags)
  If tab1.Tab < 3 Then
    'Found Something
    If Pos(tab1.Tab) >= 0 Then
    'Didn't Find
    Else
      Do Until tab1.Tab = 3 Or Pos(tab1.Tab) >= 0
        tab1.Tab = tab1.Tab + 1
        Pos(tab1.Tab) = rtb(tab1.Tab).Find(txtF.Text, Pos(tab1.Tab) + 1, , FindFlags)
      Loop
    End If
  End If
End If
End Sub
Private Sub rtb_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
rtb(tab1.Tab).SetFocus
If rtb(tab1.Tab).SelLength > 0 Then
  cmdCase(0).Enabled = True
  cmdCase(1).Enabled = True
  mnuCopy.Enabled = True
  mnuCut.Enabled = True
  mnuDelete.Enabled = True
  mnuPrintSelected.Enabled = True
Else
  cmdCase(0).Enabled = False
  cmdCase(1).Enabled = False
  mnuCopy.Enabled = False
  mnuCut.Enabled = False
  mnuDelete.Enabled = False
  mnuPrintSelected.Enabled = False
End If
End Sub
Private Sub rtb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, y As Single)
Dim i As Integer
On Error Resume Next
'Fixes bug
For i = 0 To 3
  Pos(i) = 0
Next
Pos(tab1.Tab) = rtb(tab1.Tab).SelStart
End Sub
Private Sub tab1_Click(PreviousTab As Integer)
Dim i As Integer
'Fixes bug
For i = 0 To 3
  Pos(i) = -1
Next
Me.Caption = "MultiDoc     " & OpenFile(tab1.Tab)
End Sub
Public Sub AdjustButtons()
Dim i(3) As Boolean, z As Integer
Dim a(3) As Integer
i(0) = cd1.FontBold
i(1) = cd1.FontItalic
i(2) = cd1.FontUnderline
i(3) = cd1.FontStrikethru
For z = 0 To 3
  If i(z) Then
    a(z) = 1
  ElseIf i(z) = False Then
    a(z) = 0
  End If
Next
'Bold
chk(1).Value = a(0)
'Italic 2
chk(2).Value = a(1)
'Underline 3
chk(3).Value = a(2)
'StrikeThru 5
chk(5).Value = a(3)
End Sub
Private Sub vsbSize_Change()
lblSize.Caption = vsbSize.Value
rtb(tab1.Tab).SetFocus
End Sub
Private Sub vsbSize_Scroll()
lblSize.Caption = vsbSize.Value
rtb(tab1.Tab).SetFocus
End Sub
Public Sub ResetButtons()
'Size
vsbSize.Value = 10
'Left Align
optAlign(0).Value = True
'Bullet
chk(0).Value = 0
'Bold
chk(1).Value = 0
'Italic 2
chk(2).Value = 0
'Underline 3
chk(3).Value = 0
'StrikeThru 5
chk(5).Value = 0
'Color
chk(4).Value = 0
'Font Name
FName = txtF.FontName
rtb(tab1.Tab).SelFontName = FName
lblFont.Caption = FName
End Sub
Public Sub BetterFind(FindText As String)
'Finds Text Correctly!!!!!
ReplaceYes = True
Dim X, y, z
big$ = rtb(tab1.Tab).Text
target$ = FindText
Do
  X = y + 1
  z = z + 1
  y = InStr(X, big$, target$)
Loop Until y = 0
If MsgBox(target$ & " was found " & z - 1 & " times.  Replace All?", vbQuestion + 1, "Replace All?") = vbCancel Then
  ReplaceYes = False
End If
End Sub
Public Sub ReplaceAll()
On Error Resume Next
OpenFix = True
'Searches & Replaces
rtb(tab1.Tab).SelText = txtR.Text
SearchMe
'Found Something
If Pos(tab1.Tab) >= 0 Then
  ReplaceAll
'Didn't Find
Else
  MsgBox "String not found"
  cmdR.Enabled = False
  cmdRAll.Enabled = False
End If
OpenFix = False
End Sub
'Extracts Text using a token or character
Public Function ExtractText(FullText As String, token As String, Optional StartAtLeft = True, Optional IncludeLeftSide = True) As String
Dim i As Integer
If StartAtLeft And IncludeLeftSide Then
  ExtractText = FullText
  For i = 1 To Len(FullText)
    If Mid(FullText, i, 1) = token Then
      ExtractText = Left(FullText, i - 1)
      Exit Function
    End If
  Next
ElseIf StartAtLeft And IncludeLeftSide = False Then
  ExtractText = FullText
  For i = 1 To Len(FullText)
    If Mid(FullText, i, 1) = token Then
      ExtractText = Right(FullText, Len(FullText) - i)
      Exit Function
    End If
  Next
ElseIf StartAtLeft = False And IncludeLeftSide Then
  ExtractText = ""
  For i = Len(FullText) To 1 Step -1
    If Mid(FullText, i, 1) = token Then
      ExtractText = Left(FullText, i - 1)
      Exit Function
    End If
  Next
ElseIf StartAtLeft = False And IncludeLeftSide = False Then
  ExtractText = ""
  For i = Len(FullText) To 1 Step -1
    If Mid(FullText, i, 1) = token Then
      ExtractText = Right(FullText, Len(FullText) - i)
      Exit Function
    End If
  Next
End If
End Function
