Attribute VB_Name = "ScreenRes"
Option Explicit

Private Const CCDEVICENAME = 32
Private Const CCFORMNAME = 32

Private Const DISP_CHANGE_SUCCESSFUL = 0
Private Const DISP_CHANGE_RESTART = 1
Private Const DISP_CHANGE_FAILED = -1
Private Const DISP_CHANGE_BADMODE = -2
Private Const DISP_CHANGE_NOTUPDATED = -3
Private Const DISP_CHANGE_BADFLAGS = -4
Private Const DISP_CHANGE_BADPARAM = -5

Private Const CDS_UPDATEREGISTRY = &H1
Private Const CDS_TEST = &H2

Private Const DM_BITSPERPEL = &H40000
Private Const DM_PELSWIDTH = &H80000
Private Const DM_PELSHEIGHT = &H100000

Private Type DEVMODE
  dmDeviceName As String * CCDEVICENAME
  dmSpecVersion As Integer
  dmDriverVersion As Integer
  dmSize As Integer
  dmDriverExtra As Integer
  dmFields As Long
  dmOrientation As Integer
  dmPaperSize As Integer
  dmPaperLength As Integer
  dmPaperWidth As Integer
  dmScale As Integer
  dmCopies As Integer
  dmDefaultSource As Integer
  dmPrintQuality As Integer
  dmColor As Integer
  dmDuplex As Integer
  dmYResolution As Integer
  dmTTOption As Integer
  dmCollate As Integer
  dmFormName As String * CCFORMNAME
  dmUnusedPadding As Integer
  dmBitsPerPel As Integer
  dmPelsWidth As Long
  dmPelsHeight As Long
  dmDisplayFlags As Long
  dmDisplayFrequency As Long
End Type

Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long

Function ChangeScreenSettings(lWidth As Integer, lHeight As Integer, lColors As Integer)
Dim tDevMode As DEVMODE, lTemp As Long, lIndex As Long
lIndex = 0
Do
  lTemp = EnumDisplaySettings(0&, lIndex, tDevMode)
  If lTemp = 0 Then Exit Do
  lIndex = lIndex + 1

  With tDevMode
    If .dmPelsWidth = lWidth And .dmPelsHeight = lHeight And .dmBitsPerPel = lColors Then
      lTemp = ChangeDisplaySettings(tDevMode, CDS_UPDATEREGISTRY)
      Exit Do
    End If
  End With
Loop
Select Case lTemp
  Case DISP_CHANGE_SUCCESSFUL
    'MsgBox "Succes!", vbInformation
  Case DISP_CHANGE_RESTART
    MsgBox "Se necesita reiniciar el equipo!", vbInformation
  Case DISP_CHANGE_FAILED
    MsgBox "Resolucion no válida", vbCritical
  Case DISP_CHANGE_BADMODE
    MsgBox "Resolucion no válida", vbCritical
  Case DISP_CHANGE_NOTUPDATED
    MsgBox "Error de registro", vbCritical
  Case DISP_CHANGE_BADFLAGS
    MsgBox "Resolucion no válida", vbCritical
End Select
End Function

