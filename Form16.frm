VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form16 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2940
   Icon            =   "Form16.frx":0000
   LinkTopic       =   "Form16"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2715
   ScaleWidth      =   2940
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   30
      Left            =   120
      ScaleHeight     =   30
      ScaleWidth      =   2700
      TabIndex        =   1
      Top             =   2520
      Width           =   2700
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000014&
         BorderStyle     =   0  'None
         Height          =   10
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   8895
         TabIndex        =   2
         Top             =   10
         Width           =   8895
      End
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   56623105
      TitleBackColor  =   33023
      TitleForeColor  =   16777215
      CurrentDate     =   38064
   End
End
Attribute VB_Name = "Form16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MonthView1.Value = Date
End Sub


