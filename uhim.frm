VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "������ OPC HDA � �������� ������ � MS SQL v1.00           (��� ""��������������"")"
   ClientHeight    =   6240
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   10260
   Icon            =   "uhim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6240
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnuFile 
      Caption         =   "����"
      Begin VB.Menu mnuSave 
         Caption         =   "���������"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "�������"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnu 
         Caption         =   ""
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuExit 
         Caption         =   "�����"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuServer 
      Caption         =   "������"
      Begin VB.Menu mnuSetup 
         Caption         =   "���������"
      End
      Begin VB.Menu mnuTake 
         Caption         =   "�������"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "������"
      Begin VB.Menu mnuAbout 
         Caption         =   "� ���������..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ModeRecieve As Integer ' ���� 0 - ����������, 1 - �����������
Public Duration As Long ' ������� ������
Public dStart As Date
Public dEnd As Date
Public iBounds As Integer
Public lValues As Long

Private Sub Form_Load()
Dim sRecieve As String
sRecieve = GetINI(App.Path & "\uhim.ini", "�����", "Mode", "0")
ModeRecieve = Val(sRecieve)
sRecieve = GetINI(App.Path & "\uhim.ini", "�����", "Duration", "500")
Duration = Val(sRecieve)
If ModeRecieve = 1 Then
    Dialog.Label2.Visible = True
Else
    Dialog.Label2.Visible = False
End If
'frmOptions.DTPicker2.Value = GetINI(App.Path & "\uhim.ini", "������", "d2", "0")
'frmOptions.DTPicker1.Value = GetINI(App.Path & "\uhim.ini", "������", "d1", "0")
dEnd = GetINI(App.Path & "\uhim.ini", "������", "d2", "0")
dStart = GetINI(App.Path & "\uhim.ini", "������", "d1", "0")
iBounds = GetINI(App.Path & "\uhim.ini", "�����������", "����", "0")
lValues = GetINI(App.Path & "\uhim.ini", "�����", "Number", "0")
End Sub

Private Sub Form_Unload(Cancel As Integer)
If MsgBox("�� ������������� ������ �����?", vbQuestion + vbYesNo) = vbNo Then
    Cancel = 1
Else
    End
End If
End Sub

Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuDoit_Click()
frmODBC.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuSetup_Click()
frmOptions.Show
End Sub

Private Sub mnuTake_Click()
Dialog.Show 1, Me
End Sub
