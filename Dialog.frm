VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Dialog 
   Caption         =   "������ OPC HDA � �������� ������ � ��   v1.00a           (��� ""��������������"")"
   ClientHeight    =   9690
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   11130
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9690
   ScaleWidth      =   11130
   Begin VB.Frame Frame6 
      Caption         =   "��������� "
      Height          =   4305
      Left            =   6330
      TabIndex        =   13
      Top             =   390
      Width           =   4635
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3840
         Top             =   2460
      End
      Begin VB.TextBox Text7 
         Height          =   345
         Left            =   930
         TabIndex        =   33
         Top             =   3180
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Height          =   315
         Left            =   930
         TabIndex        =   32
         Top             =   2820
         Width           =   2055
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "���������"
         Height          =   345
         Left            =   3330
         TabIndex        =   31
         Top             =   3900
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   3720
         TabIndex        =   21
         Text            =   "1"
         Top             =   810
         Width           =   825
      End
      Begin VB.CheckBox Check1 
         Caption         =   "�����������"
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   330
         Width           =   1485
      End
      Begin VB.Frame Frame1 
         Caption         =   "�����"
         Height          =   1095
         Left            =   90
         TabIndex        =   17
         Top             =   210
         Width           =   2535
         Begin VB.OptionButton Option2 
            Caption         =   "����������� ���������"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   600
            Value           =   -1  'True
            Width           =   2295
         End
         Begin VB.OptionButton Option1 
            Caption         =   "���������� ���������"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   930
         TabIndex        =   16
         Text            =   " "
         Top             =   3600
         Width           =   2055
      End
      Begin VB.TextBox Text4 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   930
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   3930
         Width           =   2055
      End
      Begin VB.TextBox Text5 
         Height          =   345
         Left            =   930
         TabIndex        =   14
         Top             =   2400
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2550
         TabIndex        =   22
         Top             =   1560
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yy HH:mm:ss"
         Format          =   52428803
         CurrentDate     =   40721.0416666667
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   150
         TabIndex        =   23
         Top             =   1560
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd-MM-yy HH:mm:ss"
         Format          =   52428803
         CurrentDate     =   40721.0416666667
         MinDate         =   -109204
      End
      Begin VB.Label Label8 
         Caption         =   "�������"
         Height          =   225
         Left            =   120
         TabIndex        =   35
         Top             =   3270
         Width           =   705
      End
      Begin VB.Label Label7 
         Caption         =   "��"
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   2880
         Width           =   765
      End
      Begin VB.Label Label4 
         Caption         =   "���-�� �����"
         Height          =   225
         Left            =   2670
         TabIndex        =   29
         Top             =   870
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "����� ������� ������"
         Height          =   195
         Left            =   2550
         TabIndex        =   28
         Top             =   1350
         Width           =   1755
      End
      Begin VB.Label Label2 
         Caption         =   "������ ������� ������"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   1350
         Width           =   1995
      End
      Begin VB.Label Label1 
         Caption         =   "���"
         Height          =   255
         Left            =   150
         TabIndex        =   26
         Top             =   3630
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "������"
         Height          =   255
         Left            =   150
         TabIndex        =   25
         Top             =   3990
         Width           =   825
      End
      Begin VB.Label Label6 
         Caption         =   "������"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   2430
         Width           =   765
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "���������"
      Height          =   4065
      Left            =   150
      TabIndex        =   11
      Top             =   5520
      Width           =   6195
      Begin VB.CommandButton cmdRem 
         Caption         =   "���������"
         Height          =   375
         Left            =   4770
         TabIndex        =   36
         Top             =   3630
         Width           =   1365
      End
      Begin VB.ListBox List4 
         Enabled         =   0   'False
         Height          =   3375
         Left            =   90
         TabIndex        =   12
         Top             =   240
         Width           =   6075
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "��� HDA �������"
      Height          =   1095
      Left            =   150
      TabIndex        =   7
      Top             =   390
      Width           =   6195
      Begin VB.CheckBox Check2 
         Caption         =   "���������"
         Height          =   375
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   240
         Width           =   945
      End
      Begin VB.CommandButton OKButton 
         Caption         =   "������"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2370
         TabIndex        =   38
         ToolTipText     =   "�������� ������ ����������� ������ �� OPC HDA"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "������������"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1020
         TabIndex        =   37
         ToolTipText     =   "����������� � OPC HDA �������"
         Top             =   240
         Width           =   1365
      End
      Begin VB.CommandButton Command3 
         Height          =   795
         Left            =   5310
         Picture         =   "Dialog.frx":548A
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   180
         Width           =   825
      End
      Begin VB.ComboBox ListBox 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   690
         Width           =   3615
      End
      Begin VB.CommandButton Command2 
         Caption         =   "��������"
         Height          =   375
         Left            =   60
         TabIndex        =   9
         ToolTipText     =   "������������ ������ OPC HDA ��������"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "�����"
         Height          =   375
         Left            =   4320
         TabIndex        =   8
         Top             =   240
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   3750
         Shape           =   3  'Circle
         Top             =   750
         Width           =   285
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   3750
         Shape           =   3  'Circle
         Top             =   690
         Width           =   315
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
      Height          =   4005
      Left            =   150
      TabIndex        =   3
      Top             =   1500
      Width           =   6195
      Begin VB.ListBox List1 
         Height          =   3570
         Left            =   90
         TabIndex        =   6
         Top             =   360
         Width           =   1935
      End
      Begin VB.ListBox List2 
         Height          =   3570
         Left            =   2010
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
      Begin VB.ListBox List3 
         Height          =   3570
         Left            =   4050
         TabIndex        =   4
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "������"
      Height          =   4875
      Left            =   6330
      TabIndex        =   1
      Top             =   4710
      Width           =   4665
      Begin VB.TextBox Text1 
         Height          =   4575
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   4545
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   9645
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11085
      _ExtentX        =   19553
      _ExtentY        =   17013
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�������"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Dialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Base 1
Dim Duration As Long ' ������� ������
Dim dStart As Date '��������� ���� �������
Dim dEnd As Date '�������� ���� �������
Dim iBounds As Integer ' �����������
Dim lValues As Long ' ����� ��������
Dim ModeRecieve As Integer ' ����� ������ - �����., �������.
'
Dim AllOPCHDAServers As Variant ' ��� �������� HDA
Dim WithEvents AnOPCHDAServer As OPCHDAServer ' ������ HDA ��� ������
Attribute AnOPCHDAServer.VB_VarHelpID = -1
Dim WithEvents AnOPCHDAItemCollection As OPCHDAItems ' ��������� ���� ������
Attribute AnOPCHDAItemCollection.VB_VarHelpID = -1
Dim AnOPCHDAServerBrowser As OPCHDABrowser ' ���������� ������� ��� ���������
Dim adoCon As ADODB.Connection
Dim rs As ADODB.Recordset

Private sText As String

' ������� ����������� ��� ����������� ������
Private Sub AnOPCHDAItemCollection_AsyncReadComplete(ByVal TransactionID As Long, ByVal Status As Long, _
                                ByVal NumItems As Long, ClientHandles() As Long, Aggregates() As Long, _
                                ItemValues() As Variant, Errors() As Long)
Dim i As Integer, j As Integer
' write your client code here to process the data values
For i = 1 To NumItems
    ' ����� ��� ���������
    Text1.Text = Text1.Text & AnOPCHDAItemCollection.Item(i).ItemID & vbCrLf
     '������� ��� ���������� ��������
    For j = 1 To ItemValues(i).Count
        'Text1.Text = Text1.Text & ItemValues(i).Item(j) & vbCrLf
        ' ������������ ������ ������ - ����� �����.���-���, ��� ��������
        Text1.Text = Text1.Text & ItemValues(i).Item(j).TimeStamp & " - " & _
            IIf(IsEmpty(ItemValues(i).Item(j).DataValue), "Empty", ItemValues(i).Item(j).DataValue) & vbCrLf
        rs.AddNew ' �������� �������� � �� MSSQL
        rs.Fields(0) = AnOPCHDAItemCollection.Item(i).ItemID '��� ���������
        rs.Fields(1) = ItemValues(i).Item(j).TimeStamp '�����
        rs.Fields(2) = ItemValues(i).Item(j).DataValue ' ��������
        rs.Update '��������� � ��
   Next j
Next i
'Show 1, Me
End Sub

Private Sub CancelButton_Click()
'AnOPCHDAServer.Disconnect ' ������������� �� ������� HDA
Unload Me ' ���������
End Sub

'Private Sub cmdPered_Click()
'
' '  SQL Server ODBC Driver
' Set adoCon = New ADODB.Connection
'adoCon.Open "Provider=SQLOLEDB;data Source=" & Me.Text5.Text & ";" & _
'"Initial Catalog=" & Me.Text6.Text & ";User Id=" & Me.Text3.Text & ";Password=" & Me.Text4.Text
'Set rs = New ADODB.Recordset
'rs.Open "Select * from " & Me.Text7.Text, adoCon, adOpenKeyset, adLockOptimistic
'rs.AddNew
'rs.Fields(0) = Now
'rs.Fields(1) = 3.1415629
'rs.Update
'rs.Close
'adoCon.Close
'Set rs = Nothing
'Set adoCon = Nothing
'
'End Sub

Private Sub Check2_Click()
    ' ������� ������ ��������������� ���������� ��������
    Call WriteINI(App.Path & "\uhim.ini", "Timer", "Start", Me.Check2.Value)
End Sub
'���������� ������ ���������
Private Sub cmdRem_Click()
Dim i As Integer
' ���������� �������� ������� � OPC HDA �������
Call WriteINI(App.Path & "\uhim.ini", "Apparats", "HDAServer", ListBox.List(ListBox.ListIndex))
Call WriteINI(App.Path & "\uhim.ini", "Apparats", "Count", List4.ListCount) ' ���-�� ����������
For i = 0 To List4.ListCount - 1 '����� ���-����
    Call WriteINI(App.Path & "\uhim.ini", "Apparats", Trim(Str(i + 1)), List4.List(i))
Next
End Sub
' ���������� ������ ���������
Private Sub cmdSave_Click()

If Me.Option1.Value Then
    Call WriteINI(App.Path & "\uhim.ini", "�����", "Mode", "0") ' ������. �����.����� ������
    ModeRecieve = 0
End If
If Me.Option2.Value Then
    Call WriteINI(App.Path & "\uhim.ini", "�����", "Mode", "1") ' ������. ����.����� ������
    ModeRecieve = 1
End If
'Call WriteINI(App.Path & "\uhim.ini", "�����", "Duration", Me.Text1.Text)
'Duration = Val(Me.Text1.Text)
dStart = Me.DTPicker1.Value ' ����. ����.���.����
dEnd = Me.DTPicker2.Value ' ����. ����.�����.����
Call WriteINI(App.Path & "\uhim.ini", "������", "d1", Trim(Str(dStart)))
Call WriteINI(App.Path & "\uhim.ini", "������", "d2", Trim(Str(dEnd)))
Call WriteINI(App.Path & "\uhim.ini", "�����", "Number", Me.Text2.Text) ' ����.����� �������� ��� ������
lValues = Me.Text2.Text
If Me.Check1.Value Then ' �������
    Call WriteINI(App.Path & "\uhim.ini", "�����������", "����", "1")
    iBounds = 1
Else
    Call WriteINI(App.Path & "\uhim.ini", "�����������", "����", "0")
    iBounds = 0
End If

Call WriteINI(App.Path & "\uhim.ini", "��", "Server", Me.Text5.Text) ' ��� ������� ��
Call WriteINI(App.Path & "\uhim.ini", "��", "BD", Me.Text6.Text) ' ��� ��
Call WriteINI(App.Path & "\uhim.ini", "��", "Table", Me.Text7.Text) ' ��� �������
Call WriteINI(App.Path & "\uhim.ini", "��", "User", Me.Text3.Text) ' ������������
Call WriteINI(App.Path & "\uhim.ini", "��", "Pass", Me.Text4.Text) ' ������



MsgBox "���������"
'Unload Me
End Sub



' ������������ �\� ���������
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    'handle ctrl+tab to move to the next tab
    If Shift = vbCtrlMask And KeyCode = vbKeyTab Then
        i = TabStrip1.SelectedItem.Index
        If i = TabStrip1.Tabs.Count Then
            'last tab so we need to wrap to tab 1
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(1)
        Else
            'increment the tab
            Set TabStrip1.SelectedItem = TabStrip1.Tabs(i + 1)
        End If
    End If
End Sub



Private Sub Option1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Option2.Value = Not Me.Option1.Value '������� ������ � �����->����
'Dialog.Label2.Visible = False

End Sub

Private Sub Option2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Option1.Value = Not Me.Option2.Value '������� ������ � ����->�����
'Dialog.Label2.Visible = True

End Sub

'Private Sub TabStrip1_Click()
'
'
'    Dim i As Integer
'    'show and enable the selected tab's controls
'    'and hide and disable all others
'    For i = 0 To tbsOptions.Tabs.Count - 1
'        If i = tbsOptions.SelectedItem.Index - 1 Then
'            picOptions(i).Left = 210
'            picOptions(i).Enabled = True
'        Else
'            picOptions(i).Left = -20000
'            picOptions(i).Enabled = False
'        End If
'    Next
'
'End Sub


'Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
'sText = Me.Text1.Text
'End Sub
'
'Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode < 48 Or KeyCode > 57 Then
'    If KeyCode <> 8 And KeyCode <> 13 And KeyCode <> 46 Then Me.Text1.Text = sText
'End If
'End Sub
' ����������� � OPCHDA �������
Private Sub Command1_Click()
Dim Branches() As String
Dim i As Integer
On Error GoTo errConnect
'MsgBox ListBox.List(ListBox.ListIndex)
' ��������� ����������� � ������� HDA
AnOPCHDAServer.Connect (AllOPCHDAServers(ListBox.ListIndex + 1))
Set AnOPCHDAServerBrowser = AnOPCHDAServer.CreateBrowser() ' �������� �������� ������ �����
Branches = AnOPCHDAServerBrowser.OPCHDABranches ' ������� ��� ���� 1 �� ������
For i = LBound(Branches) To UBound(Branches)
    List1.AddItem Branches(i)
Next i
' �������� ������ ������
'Set AnOPCHDAItemCollection = AnOPCHDAServer.OPCHDAItems
'AnOPCHDAItemCollection.AddItem "Random.Real4", 1
'AsyncReadRawButton 'SyncReadRaw
AnOPCHDAItemCollection.RemoveAll ' �������� ���������� ���-��
Me.Shape1.FillColor = RGB(0, 255, 0) ' ������ ������
errConnect:
End Sub
' �������� ������ OPC HDA ��������
Private Sub Command2_Click()
'������� ������ hda ��������
Dim i As Integer

'Set AnOPCHDAServer = New OPCHDAServer ' ��������� ������ HDA
' ���������� ������ �������� HDA
AllOPCHDAServers = AnOPCHDAServer.GetOPCHDAServers
For i = LBound(AllOPCHDAServers) To UBound(AllOPCHDAServers)
    ListBox.AddItem AllOPCHDAServers(i)
Next i
Me.Command1.Enabled = True
Me.ListBox.Enabled = True
Me.Shape1.FillColor = RGB(255, 255, 0) '������ ���������� ������
End Sub




Private Sub Command3_Click()
frmAbout.Show 1, Me ' ������ �������� ���������
End Sub

Private Sub Form_Load()
' ���������� �������� ��� �������� ���������
Dim sRecieve As String
    'center the form
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
sRecieve = GetINI(App.Path & "\uhim.ini", "�����", "Mode", "0") '����� ������
ModeRecieve = Val(sRecieve)
sRecieve = GetINI(App.Path & "\uhim.ini", "�����", "Duration", "500") ' ������������ - ?
Duration = Val(sRecieve)
'frmOptions.DTPicker2.Value = GetINI(App.Path & "\uhim.ini", "������", "d2", "0")
'frmOptions.DTPicker1.Value = GetINI(App.Path & "\uhim.ini", "������", "d1", "0")
dEnd = GetINI(App.Path & "\uhim.ini", "������", "d2", "0") ' �����.������
dStart = GetINI(App.Path & "\uhim.ini", "������", "d1", "0") ' ���.������
iBounds = GetINI(App.Path & "\uhim.ini", "�����������", "����", "0") '�����������
lValues = GetINI(App.Path & "\uhim.ini", "�����", "Number", "0") ' ����� ����������� ��������
' ���� ����������� ����������
    If ModeRecieve = 1 Then
        Me.Option1.Value = False
        Me.Option2.Value = True
    Else
        Me.Option1.Value = True
        Me.Option2.Value = False
    End If
    'Me.Text1.Text = Duration
    Me.Check1 = iBounds
    Me.Text2.Text = lValues
    Me.DTPicker2.Value = dEnd
    Me.DTPicker1.Value = dStart
' ���� ���������� ������� � ��
Text5.Text = GetINI(App.Path & "\uhim.ini", "��", "Server", "?")
Text6.Text = GetINI(App.Path & "\uhim.ini", "��", "BD", "?")
Text7.Text = GetINI(App.Path & "\uhim.ini", "��", "Table", "?")
Text3.Text = GetINI(App.Path & "\uhim.ini", "��", "User", "?")
Text4.Text = GetINI(App.Path & "\uhim.ini", "��", "Pass", "?")

Set AnOPCHDAServer = New OPCHDAServer ' ��������� ������ HDA
Set AnOPCHDAItemCollection = AnOPCHDAServer.OPCHDAItems
' ���������� � ������ ������� HDA ������������ ���������
Dim ARealOPCHDAServer As String, dCount As String, i As Integer, sNam As String
ARealOPCHDAServer = GetINI(App.Path & "\uhim.ini", "Apparats", "HDAServer", "xxx") '"Matrikon.OPC.Simulation.1"
AnOPCHDAServer.Connect (ARealOPCHDAServer)
' ����� ������� ������ �������� ������
' �������� ������ ������
dCount = GetINI(App.Path & "\uhim.ini", "Apparats", "Count", "0")
AnOPCHDAItemCollection.RemoveAll
For i = 1 To CInt(dCount)
    sNam = GetINI(App.Path & "\uhim.ini", "Apparats", Trim(Str(i)), "***")
    AnOPCHDAItemCollection.AddItem sNam, i
Next
' ���������� �� ������ �������� �������������
Me.Check2.Value = GetINI(App.Path & "\uhim.ini", "Timer", "Start", "0")

End Sub
' ������ ������� �����
Private Sub List1_Click()
Dim Items() As String
Dim i As Integer
'������� � ����� ������ �����
AnOPCHDAServerBrowser.MoveToRoot
List2.Clear ' ���������� �� ���������� ��������
List2.Refresh
List3.Clear ' ���������� �� ���������� ��������
List3.Refresh
' ������� �������� �����
AnOPCHDAServerBrowser.MoveDown (AnOPCHDAServerBrowser.OPCHDABranches(List1.ListIndex + 1))
Items = AnOPCHDAServerBrowser.OPCHDABranches ' ���� ���� �������� - �������
For i = LBound(Items) To UBound(Items)
    List2.AddItem Items(i)
Next i
Items = AnOPCHDAServerBrowser.OPCHDAItems ' ���� ���� �������� - �������
For i = LBound(Items) To UBound(Items)
    List2.AddItem Items(i)
Next i
List2.Refresh
End Sub
' ������ ������� �����
Private Sub List2_Click()
Dim Items() As String
Dim i As Integer
List3.Clear ' ���������� �� ���������� ��������
List3.Refresh
AnOPCHDAServerBrowser.MoveDown (AnOPCHDAServerBrowser.OPCHDABranches(List2.ListIndex + 1))
Items = AnOPCHDAServerBrowser.OPCHDABranches ' ���� ���� �������� - �������
For i = LBound(Items) To UBound(Items)
    List3.AddItem Items(i)
Next i
Items = AnOPCHDAServerBrowser.OPCHDAItems ' ���� ���� �������� - �������
For i = LBound(Items) To UBound(Items)
    List3.AddItem Items(i)
Next i
List3.Refresh
AnOPCHDAServerBrowser.MoveUp
End Sub
' ������ ������� �����
Private Sub List3_Click()
List4.Enabled = True
List4.AddItem List2.List(List2.ListIndex) & "." & List3.List(List3.ListIndex)
List4.Refresh
' ����� ������� ������ �������� ������
AnOPCHDAItemCollection.AddItem List4.List(List4.ListCount - 1), List4.ListCount
OKButton.Enabled = True
End Sub

' ������������ ������� ������� DELETE � ������ �����
Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
Dim errCode() As Long, iIndex As Long
Dim ServerHandles() As Long
Dim NumItems As Long
If KeyCode = vbKeyDelete Then
    NumItems = AnOPCHDAItemCollection.Count ' 1 '0 ����� �����
    iIndex = 0
    ReDim ServerHandles(NumItems) ' ������ ��������� �����
    iIndex = List4.ListIndex + 1 ' c�-�� �����
    List4.RemoveItem List4.ListIndex ' ������� ��������� ��� �� ������
    ServerHandles(1) = AnOPCHDAItemCollection.Item(iIndex).ServerHandle
    AnOPCHDAItemCollection.Remove 1, ServerHandles, errCode ' ������� ��� �� ��������� �����
    If List4.ListCount = 0 Then ' ���� ������� ��� ���� ������� ������ � ������
        List4.Enabled = False
        OKButton.Enabled = False
    End If
End If
End Sub

' ���������� ������ ������
Private Sub OKButton_Click()
Text1.Text = "" ' ������ ����������
 '  SQL Server ODBC Driver
 ' ��������� ������ � ������ MSSQL
 Set adoCon = New ADODB.Connection
adoCon.Open "Provider=SQLOLEDB;data Source=" & Me.Text5.Text & ";" & _
"Initial Catalog=" & Me.Text6.Text & ";User Id=" & Me.Text3.Text & ";Password=" & Me.Text4.Text
Set rs = New ADODB.Recordset ' ����� ������� ��� ������ ��������
rs.Open "Select * from " & Me.Text7.Text, adoCon, adOpenKeyset, adLockOptimistic
'' ������� ���������
'Dim addItemCount As Long
'Dim AnOPCHDAItemServerErrors() As Long
'Dim AnOPCHDAItemIDs() As String ' ������ ���� ������
'Dim i As Integer
'addItemCount = AnOPCHDAItemCollection.Count
'ReDim AnOPCHDAItemIDs(addItemCount)
'For i = 1 To addItemCount
'    AnOPCHDAItemIDs(i) = List4.List(i - 1)
'Next
'AnOPCHDAItemCollection.Validate addItemCount, AnOPCHDAItemIDs, AnOPCHDAItemServerErrors
'
'If List4.Enabled = True Then
 ' ��������� ������ � ����������� �� ������ ������
    If Option2.Value = True Then
        ' � ����������� ������
        AsyncReadRawButton
    Else
        ' � ���������� ������
        SyncReadRaw
        'Show 1, Me
    End If
'Else
'    MsgBox "������� ���� ��� ������"
'End If
End Sub
' ���������� �����.������
Private Sub AsyncReadRawButton()
Dim TransactionID As Long, i As Integer
Dim StartTime As Date
Dim EndTime As Date
Dim NumValues As Long
Dim Bounds As Boolean
Dim NumItems As Long
Dim ServerHandles(10) As Long
Dim Errors() As Long
Dim CancelID As Long
On Error GoTo err_AsyncRead
NumItems = AnOPCHDAItemCollection.Count '1 'List4.ListCount ���������� ����� ��� ������
For i = 1 To NumItems
    ' set up which items to be read
    'ServerHandles(i) = i  'AnOPCItemServerHandles(i)
    ServerHandles(i) = AnOPCHDAItemCollection.Item(i).ServerHandle
Next i
TransactionID = 1
'StartTime = "-1D"
'EndTime = "NOW"
'NumValues = 10
'Bounds = False
StartTime = Me.DTPicker1.Value 'dStart ������ �������
EndTime = Me.DTPicker2.Value 'dEnd ����� �������
NumValues = Me.Text2.Text 'lValues ����� ����������� ��������
Bounds = Me.Check1.Value 'IIf(iBounds = 1, True, False) ���� �� �������
' ��������������� ��������� ������
CancelID = AnOPCHDAItemCollection.AsyncReadRaw(TransactionID, StartTime, EndTime, _
                                    NumValues, Bounds, NumItems, ServerHandles, Errors)
'For i = 1 To NumItems
'    ' process errors
'    Text1.Text = Errors(i) & vbCrLf
'Next i
err_AsyncRead:
End Sub
' ���������� ����������� ������
Sub SyncReadRaw()
'Dim StartTime As Variant
'Dim EndTime As Variant
Dim StartTime As Date
Dim EndTime As Date
Dim NumValues As Long
Dim Bounds As Boolean
Dim NumItems As Long
Dim ServerHandles() As Long
Dim ItemValues() As Variant
Dim Errors() As Long, i As Long, j As Long
NumItems = AnOPCHDAItemCollection.Count ' 1 '0 ���-�� ����� ��� ������
ReDim ServerHandles(NumItems)
For i = 1 To NumItems
    ' set up which items to be read
    'ServerHandles(i) = i 'AnOPCItemServerHandles(i)
    ServerHandles(i) = AnOPCHDAItemCollection.Item(i).ServerHandle
Next i
'Dim AnOPCHDAServerTime As Date
'AnOPCHDAServerTime = AnOPCHDAServer.StartTime
'StartTime = CDate("06/27/2011 10:32:55")  'frmOptions.DTPicker1.Value  '"-1D"
'EndTime = CDate("06/27/2011 11:32:55") 'frmOptions.DTPicker2.Value '"NOW"
StartTime = Me.DTPicker1.Value 'dStart ������ �������
EndTime = Me.DTPicker2.Value 'dEnd ����� �������
NumValues = Me.Text2.Text 'lValues ����� ����������� ��������
Bounds = Me.Check1.Value 'IIf(iBounds = 1, True, False)
' ��������������� ������
AnOPCHDAItemCollection.SyncReadRaw StartTime, EndTime, NumValues, Bounds, NumItems, _
                                                        ServerHandles, ItemValues, Errors
'����� ���������� ������ ���-����
For i = 1 To NumItems
    Text1.Text = Text1.Text & AnOPCHDAItemCollection.Item(i).ItemID & vbCrLf
    For j = 1 To ItemValues(i).Count
        ' process the values
        'Text1.Text = ItemValues(i).Count & vbCrLf
        Text1.Text = Text1.Text & ItemValues(i).Item(j).TimeStamp & " - " & _
            IIf(IsEmpty(ItemValues(i).Item(j).DataValue), "Empty", ItemValues(i).Item(j).DataValue) & vbCrLf
        rs.AddNew ' ������ � ��
        rs.Fields(0) = AnOPCHDAItemCollection.Item(i).ItemID
        rs.Fields(1) = ItemValues(i).Item(j).TimeStamp
        rs.Fields(2) = ItemValues(i).Item(j).DataValue
        rs.Update
    Next j
Next i
End Sub

'Private Sub Option3_Click()
'    ' ������������ ��������������� ������ ������
'    Call WriteINI(App.Path & "\uhim.ini", "Timer", "Start", Me.Option3.Value)
'End Sub
' ���������� ������� ���� ��� �������.������
Private Sub Timer1_Timer()
If Check2.Value = 1 Then
    ' ��������� ��� � ���
    If Minute(Now()) = 0 And Second(Now()) = 0 Then
        Me.DTPicker2.Value = Now()
        Me.DTPicker1.Value = DateAdd("h", -1, Now())
        Me.Text2.Text = "1"
        OKButton_Click
    End If
End If
End Sub
