Attribute VB_Name = "mdlGeneral"
Option Explicit

' -- WIndows Directori

Public Const MAX_PATH = 260
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

' ------------ API Memories Application Checking
Declare Function FindWindow Lib "USER32" Alias _
"FindWindowA" (ByVal lpClassName As String, _
ByVal lpWindowName As String) As Long

Declare Function GetWindow Lib "USER32" (ByVal hwnd _
As Long, ByVal wCmd As Long) As Long

Declare Function OpenIcon Lib "USER32" (ByVal hwnd _
As Long) As Long

Declare Function SetForegroundWindow Lib "USER32" _
(ByVal hwnd As Long) As Long
         
Public Const GW_HWNDPREV = 3
'----------------- BARU
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetSystemMenu Lib "USER32" (ByVal hwnd _
As Long, ByVal bRevert As Boolean) As Long
   
Private Declare Function GetMenuItemCount Lib "USER32" (ByVal _
hMenu As Long) As Long
   
Private Declare Function RemoveMenu Lib "USER32" (ByVal _
hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) _
As Long
   
Private Declare Function DrawMenuBar Lib "USER32" (ByVal hwnd As Long) As Long

Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&

Public DB As Database
Public rc As Recordset

Public Function ViewFormat(ByVal Nilai As String)
    ViewFormat = Format(Nilai, "###,###,###")
End Function

Public Function DefaFormat(ByVal Nilai As String)
    DefaFormat = Format(Nilai, "#########")
End Function

Public Sub Delay(ByVal Second As Byte)
  Dim tmp As Integer, x As Integer
  tmp = Second * 10000
  Do Until x = tmp
    x = x + 2
    DoEvents
  Loop
End Sub

'' ---- Grid Procedure --------
'Public Function ReadCell(ByRef Grid As MSFlexGrid, ByVal row As Integer, ByVal col As Integer) As String
'    ReadCell = Grid.TextMatrix(row, col)
'End Function
'
'Public Sub WriteCell(ByRef Grid As MSFlexGrid, ByVal row As Integer, ByVal col As Integer, ByVal Nilai As String)
'    Grid.TextMatrix(row, col) = Nilai
'End Sub
'
'' ---- End Of Grid Procedure --------

' --------------------- Date Validation --------------------
Public Function CekTgl(ByVal tDate As String) As Boolean
 On Error GoTo akhir
 If Mid(tDate, 3, 1) <> "-" Or Mid(tDate, 6, 1) <> "-" Then GoTo akhir
 Dim Tgl As Integer, Bln As Integer, Thn As Integer
 Tgl = Left(tDate, 2)
 Bln = Mid(tDate, 4, 2)
 Thn = Right(tDate, 4)
 
 If Len(tDate) <> 10 Then
   MsgBox "Format tanggal salah, yang benar Contoh 01-01-1981", vbInformation, "Tanggal salah"
   CekTgl = False
   Exit Function
 End If
 
 If Bln > 12 Or Bln < 0 Then
   MsgBox "Maaf, tidak ada bulan " & Bln, vbInformation, "Bulan salah"
   CekTgl = False
   Exit Function
 End If
 
 If (Bln = 1 Or Bln = 3 Or Bln = 5 Or Bln = 7 Or Bln = 8 Or Bln = 12) And Tgl > 31 Then
   MsgBox "Tanggal hanya sampai tanggal 31", vbInformation, "Tanggal salah"
   CekTgl = False
   Exit Function
 ElseIf (Bln = 4 Or Bln = 6 Or Bln = 9 Or Bln = 11) And Tgl > 30 Then
   MsgBox "Tanggal hanya sampai tanggal 30", vbInformation, "Tanggal salah"
   CekTgl = False
   Exit Function
 End If
 
 If (Thn Mod 4 <> 0 And Bln = 2) And Tgl > 28 Then
   MsgBox "Maaf,tanggal pada bulan ini hanya sampai 28", vbInformation, "Tanggal salah"
   CekTgl = False
   Exit Function
 ElseIf (Thn Mod 4 = 0 And Bln = 2) And Tgl > 29 Then
   MsgBox "Maaf,tanggal pada bulan ini hanya sampai 29", vbInformation, "Tanggal salah"
   CekTgl = False
   Exit Function
 End If
 CekTgl = True
 Exit Function
akhir:
   MsgBox "Maaf tanggal salah, Contoh yang benar 01-01-1981", vbInformation, "Tanggal salah"
   CekTgl = False
End Function
' -------------- End of Date Validation ---------------------

'--------- Contol manipulation Section ------------
Public Sub DisableClose(frm As Form, Optional _
  Disable As Boolean = True)
    'Setting Disable to False disables the 'X',
     'otherwise, it's reset
    Dim hMenu As Long
    Dim nCount As Long
    If Disable Then
        hMenu = GetSystemMenu(frm.hwnd, False)
        nCount = GetMenuItemCount(hMenu)
        Call RemoveMenu(hMenu, nCount - 1, MF_REMOVE Or _
            MF_BYPOSITION)
        Call RemoveMenu(hMenu, nCount - 2, MF_REMOVE Or _
            MF_BYPOSITION)
        DrawMenuBar frm.hwnd
    Else
        GetSystemMenu frm.hwnd, True
        DrawMenuBar frm.hwnd
    End If
End Sub

Public Sub PaintControl3D(frm As Form, Ctl As Control)
    frm.Line (Ctl.Left, Ctl.Top - 15)-(Ctl.Left + _
          Ctl.Width, Ctl.Top - 15), &H808080, BF
    frm.Line (Ctl.Left - 15, Ctl.Top)-(Ctl.Left - 15, _
         Ctl.Top + Ctl.Height), &H808080, BF
    frm.Line (Ctl.Left + Ctl.Width, Ctl.Top)- _
      (Ctl.Left + Ctl.Width, Ctl.Top + Ctl.Height), &HFFFFFF, BF
    frm.Line (Ctl.Left, Ctl.Top + Ctl.Height)- _
    (Ctl.Left + Ctl.Width, Ctl.Top + Ctl.Height), &HFFFFFF, BF
End Sub

Public Sub PaintForm3D(frm As Form)
    frm.Line (0, 0)-(frm.ScaleWidth, 0), &HFFFFFF, BF
    frm.Line (0, 0)-(0, frm.ScaleHeight), &HFFFFFF, BF
    frm.Line (frm.ScaleWidth - 15, 0)-(frm.ScaleWidth - 15, _
       frm.Height), &H808080, BF
    frm.Line (0, frm.ScaleHeight - 15)-(frm.ScaleWidth, _
    frm.ScaleHeight - 15), &H808080, BF
End Sub
'--------- End of Contol manipulation Section ------------

Sub Main()
    Set DB = OpenDatabase(App.Path & "\data.mdb")
    Load frmUtama
    
    frmUtama.Show
End Sub
