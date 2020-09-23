VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example of the progressbar"
   ClientHeight    =   5310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5130
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   5130
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   3195
      Left            =   4560
      TabIndex        =   16
      Top             =   60
      Width           =   495
      Begin VB.Timer Timer5 
         Interval        =   200
         Left            =   60
         Top             =   60
      End
      Begin VB.PictureBox Picture5 
         Height          =   2955
         Left            =   60
         ScaleHeight     =   2895
         ScaleWidth      =   315
         TabIndex        =   17
         Top             =   180
         Width           =   375
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Background image (with effects)"
      Height          =   1575
      Left            =   60
      TabIndex        =   12
      Top             =   3300
      Width           =   4995
      Begin VB.Timer Timer4 
         Interval        =   100
         Left            =   3420
         Top             =   120
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   660
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1140
         Width           =   4215
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         Height          =   840
         Left            =   120
         Picture         =   "frmExample.frx":0000
         ScaleHeight     =   780
         ScaleWidth      =   4680
         TabIndex        =   13
         Top             =   240
         Width           =   4740
      End
      Begin VB.Label Label2 
         Caption         =   "Effect"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   1140
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Personal colors support"
      Height          =   795
      Left            =   60
      TabIndex        =   10
      Top             =   2460
      Width           =   4455
      Begin VB.Timer Timer3 
         Interval        =   100
         Left            =   2340
         Top             =   480
      End
      Begin VB.PictureBox Picture3 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   4095
         TabIndex        =   11
         Top             =   240
         Width           =   4155
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Example 2"
      Height          =   1155
      Left            =   60
      TabIndex        =   4
      Top             =   1260
      Width           =   4455
      Begin VB.CheckBox Check2 
         Caption         =   "Hide"
         Height          =   195
         Left            =   3180
         TabIndex        =   9
         Top             =   780
         Width           =   735
      End
      Begin VB.Timer Timer2 
         Interval        =   20
         Left            =   60
         Top             =   180
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Text            =   "Loading..."
         Top             =   720
         Width           =   1215
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   4095
         TabIndex        =   6
         Top             =   240
         Width           =   4155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start/Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Caption"
         Height          =   255
         Left            =   1140
         TabIndex        =   8
         Top             =   720
         Width           =   675
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Example 1"
      Height          =   1155
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4455
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   2160
         Top             =   660
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Use percent"
         Height          =   255
         Left            =   3060
         TabIndex        =   3
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdStart1 
         Caption         =   "Start/Stop"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   915
      End
      Begin VB.PictureBox Picture1 
         Height          =   375
         Left            =   120
         ScaleHeight     =   315
         ScaleWidth      =   4095
         TabIndex        =   1
         Top             =   240
         Width           =   4155
      End
   End
   Begin VB.Label Label3 
      Caption         =   "http://www.infotrade.it/sebamix"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   4920
      Width           =   4935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PB1 As New smProgressBar
Private PB2 As New smProgressBar
Private PB3 As New smProgressBar
Private PB4 As New smProgressBar
Private PB5 As New smProgressBar

Private Effects(1 To 12) As Long

Private Sub Check1_Click()
    If Check1.Value = vbChecked Then
        PB1.UsePercent = True
    Else
        PB1.UsePercent = False
    End If
    
End Sub

Private Sub cmdStart1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Command1_Click()
    Timer2.Enabled = Not Timer2.Enabled
End Sub

Private Sub Form_Load()
    PB1.InitPB Picture1, SMHorizontal
    PB1.Max = 150
    
    PB2.InitPB Picture2, SMHorizontal
    PB2.Max = 1000
    
    PB3.InitPB Picture3, SMHorizontal
    PB3.Max = 1000
    PB3.BackColor = &HBBBBBB
    PB3.ForeColor = vbBlack
    PB3.ProgressColor = vbYellow
    PB3.Caption = "Use your fav. colors!"
    
    PB4.InitPB Picture4, SMHorizontal
    PB4.Max = 1000
    PB4.Caption = "Your images as ProgressBar"
    PB4.ForeColor = vbBlack
    
    
    Combo1.AddItem "SM_Color_Gray_Progress"
    Effects(1) = SMDrawStyles.SM_Color_Gray_Progress
    
    Combo1.AddItem "SM_Invert"
    Effects(2) = SMDrawStyles.SM_Invert
    
    Combo1.AddItem "SM_MaskNotPen"
    Effects(3) = SMDrawStyles.SM_MaskNotPen
    
    Combo1.AddItem "SM_MaskPen"
    Effects(4) = SMDrawStyles.SM_MaskPen
    
    Combo1.AddItem "SM_MaskPenNot"
    Effects(5) = SMDrawStyles.SM_MaskPenNot
    
    Combo1.AddItem "SM_MergeNotPen"
    Effects(6) = SMDrawStyles.SM_MergeNotPen
    
    Combo1.AddItem "SM_MergePen"
    Effects(7) = SMDrawStyles.SM_MergePen
    
    Combo1.AddItem "SM_MergePenNot"
    Effects(8) = SMDrawStyles.SM_MergePenNot
    
    Combo1.AddItem "SM_NotMaskPen"
    Effects(9) = SMDrawStyles.SM_NotMaskPen
    
    Combo1.AddItem "SM_NotMergePen"
    Effects(10) = SMDrawStyles.SM_NotMergePen
    
    Combo1.AddItem "SM_NotXorPen"
    Effects(11) = SMDrawStyles.SM_NotXorPen
    
    Combo1.AddItem "SM_XorPen"
    Effects(12) = SMDrawStyles.SM_XorPen
    Combo1.ListIndex = 0
    
    PB5.InitPB Picture5, SMVertical
    PB5.ShowCaption = False
    PB5.ProgressColor = vbBlue
    PB5.GiveOfficeBorder
End Sub

Private Sub Timer1_Timer()
    If PB1.Value < PB1.Max Then
        PB1.Value = PB1.Value + 1
    Else
        PB1.Value = PB1.Min
    End If
End Sub

Private Sub Timer2_Timer()
    If Check2.Value = vbUnchecked Then
        PB2.ShowCaption = True
    Else
        PB2.ShowCaption = False
    End If
    PB2.Caption = Text1.Text
    If PB2.Value < PB2.Max Then
        PB2.Value = PB2.Value + 1
    Else
        PB2.Value = PB2.Min
    End If
End Sub

Private Sub Timer3_Timer()
    If PB3.Value < PB3.Max Then
        PB3.Value = PB3.Value + 1
    Else
        PB3.Value = PB3.Min
    End If
End Sub

Private Sub Timer4_Timer()
    PB4.PictureProgress = Effects(Combo1.ListIndex + 1)
    If PB4.Value < PB4.Max Then
        PB4.Value = PB4.Value + 1
    Else
        PB4.Value = PB4.Min
    End If

End Sub

Private Sub Timer5_Timer()
    If PB5.Value < PB5.Max Then
        PB5.Value = PB5.Value + 1
    Else
        PB5.Value = PB5.Min
    End If
End Sub
