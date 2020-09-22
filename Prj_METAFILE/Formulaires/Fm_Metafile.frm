VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Fm_Metafile 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "enhMetafile objects' colour editing ..."
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic_RemainingColours 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   7
      Top             =   5160
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_RemainingColours 
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   8
      Top             =   5160
      Width           =   375
   End
   Begin VB.OptionButton Opt_Remaining 
      Caption         =   "Change remaining colours to"
      Height          =   615
      Index           =   1
      Left            =   7680
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.OptionButton Opt_Remaining 
      Caption         =   "Leave remaining colours as is"
      Height          =   615
      Index           =   0
      Left            =   7680
      TabIndex        =   5
      Top             =   3840
      Value           =   -1  'True
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Load 
      Caption         =   "&Select EMF file"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Cmq_Quit 
      Caption         =   "&Quit"
      Height          =   375
      Left            =   10320
      TabIndex        =   13
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_PickColour 
      Caption         =   "C"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   3
      Top             =   2520
      Width           =   375
   End
   Begin VB.PictureBox Pic_InitColour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      ScaleHeight     =   345
      ScaleWidth      =   1425
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Validate 
      Caption         =   "Validate"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "S&ave new EMF"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   6360
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Pic_FinalColour 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7680
      ScaleHeight     =   345
      ScaleWidth      =   1065
      TabIndex        =   2
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Cmd_Display 
      Caption         =   "&Display changes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7680
      TabIndex        =   9
      Top             =   5760
      Width           =   1455
   End
   Begin VB.PictureBox Pic_Metafile 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   5415
      Left            =   120
      ScaleHeight     =   5385
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   1440
      Width           =   7335
   End
   Begin VB.ListBox Lst_Modifications 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5310
      Left            =   9240
      TabIndex        =   11
      Top             =   1440
      Width           =   2535
   End
   Begin VB.Line Line4 
      X1              =   7560
      X2              =   11880
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Line Line3 
      X1              =   7560
      X2              =   9240
      Y1              =   5640
      Y2              =   5640
   End
   Begin VB.Line Line2 
      X1              =   7560
      X2              =   9240
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line1 
      X1              =   7560
      X2              =   7560
      Y1              =   1080
      Y2              =   6840
   End
   Begin VB.Label Label4 
      Caption         =   "by this one"
      Height          =   255
      Index           =   1
      Left            =   7680
      TabIndex        =   19
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Replace this colour"
      Height          =   255
      Index           =   0
      Left            =   7680
      TabIndex        =   18
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Replace ""initial colour"";""by this colour"""
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   7560
      TabIndex        =   17
      Top             =   1080
      Width           =   4215
   End
   Begin VB.Label Lbl_PicTitle 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selected enhMetafile"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   375
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "This tutorial will show you how to easily change the objects' colours in an enhanced metafile (EMF)."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1080
      TabIndex        =   15
      Top             =   480
      Width           =   10575
   End
   Begin VB.Label Lbl_Title 
      BackStyle       =   0  'Transparent
      Caption         =   "Objects' colour editing in an EMF file."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   14
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   750
      Left            =   240
      Picture         =   "Fm_Metafile.frx":0000
      Top             =   120
      Width           =   750
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Fm_Metafile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'
'___________________________________________________________________________
' Program name      : EdCol_enhMetafile.
' Description       : A simple way to edit the object's colours in an enhanced
'                     metafile (EMF).
' Company           : MELANTECH
' Authors           : Weitten Pascal
'___________________________________________________________________________
'
' Date              : (c) 2005.05.10
' Version N°        : V0.1
' Customer          : Internal stuff.
'
' Last Modification : 2005.05.10
'___________________________________________________________________________
' TODO :
'       -
'       -
'___________________________________________________________________________
'
' By extension you should be able to modify the object and reorganize
' the complete file.
'___________________________________________________________________________
'
' How it works:
'   1- Open an EMF file.
'   2- Move your mouse around the objects click to pick th colour
'      you want to change.
'   3- Pick a colour for replacement (it feeds the list). Continue
'      the steps 2 and 3 as many times and on as many objects as you wish.
'   4- Select if the other objects' shall remain their colours or replace
'      them by a unique colour.
'   5- Display the results.
'   6- Save the results to an EMF file. Objects' will be kept only colours
'      are changed. Use i.e. Illustrator to see them or reload the new generated
'      EMF file.
'___________________________________________________________________________
'
Dim picRECT As RECT
Dim hMF As Long, ret As Long
Dim strEMF_FileName As String

Private Sub Cmd_Load_Click()
    Dim i As Integer
    
    On Error Resume Next
    On Error GoTo Err_Handler
    With Fm_Metafile
        'Cancel set to true
        .CommonDialog1.CancelError = True
        .CommonDialog1.InitDir = App.Path + "\Ressources\"
        .CommonDialog1.DialogTitle = "EMF file to display..."
        .CommonDialog1.Filter = "Enhanced Metafile (*.emf)|*.emf"
        .CommonDialog1.ShowOpen
        'Le code se situe dans le module Md_BaseDonnées
        strEMF_FileName = .CommonDialog1.FileName
        Pic_Metafile.Picture = LoadPicture(strEMF_FileName)
        Cmd_PickColour.Enabled = True
        Cmd_Validate.Enabled = True
        Cmd_Display.Enabled = True
    End With
    Exit Sub
Err_Handler:
End Sub

Private Sub Cmd_PickColour_Click()
    With CommonDialog1
        .ShowColor
        Pic_FinalColour.BackColor = .Color
    End With
End Sub

Private Sub Cmd_Validate_Click()
    Lst_Modifications.AddItem Pic_InitColour.BackColor & ";" & Pic_FinalColour.BackColor
End Sub

Private Sub Cmd_Display_Click()
    Dim Nb_Lines As Integer
    Dim i As Integer
    Dim ret As Long
    
    ret = GetClientRect(Pic_Metafile.hwnd, picRECT)
    Nb_Lines = Lst_Modifications.ListCount
    If Nb_Lines > 0 Then
        ReDim Temp_Array(Nb_Lines) As String
        For i = 0 To Nb_Lines - 1
            Let Temp_Array(i) = Lst_Modifications.List(i)
        Next i
        Call EnumEnhMetaFile(ByVal Pic_Metafile.hdc, ByVal Pic_Metafile.Picture, AddressOf EnhMetaFileProc, 0, picRECT)
        Pic_Metafile.Refresh
        Erase Temp_Array
    End If
    Cmd_Save.Enabled = True
End Sub

Private Sub Cmd_Save_Click()
    SavePicture Pic_Metafile, App.Path + "\Temp\Test.emf"
    MsgBox "File saved to: " + App.Path + "\Temp\Test.emf", vbInformation + vbOKOnly, "Statement..."
End Sub

Private Sub Cmq_Quit_Click()
    Unload Me
End Sub

Private Sub Cmd_RemainingColours_Click()
    With CommonDialog1
        .ShowColor
        Pic_RemainingColours.BackColor = .Color
    End With
End Sub

Private Sub Form_Load()
    Pic_FinalColour.BackColor = 65535   'By default assign Yellow as replacement colour.
End Sub

Private Sub Opt_Remaining_Click(Index As Integer)
    If Opt_Remaining(0).Value = True Then
        Cmd_RemainingColours.Enabled = False
    ElseIf Opt_Remaining(1).Value = True Then
        Cmd_RemainingColours.Enabled = True
    End If
End Sub

Private Sub Pic_Metafile_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Pic_InitColour.BackColor = Pic_Metafile.Point(x, Y)
End Sub
