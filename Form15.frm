VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Form15 
   Caption         =   "Calendario"
   ClientHeight    =   2460
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2610
   LinkTopic       =   "Form15"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   2610
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   51052545
      TitleBackColor  =   33023
      TitleForeColor  =   16777215
      CurrentDate     =   38064
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    MonthView1.Value = Date
End Sub
