VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DBF"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7185
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdNext 
      Caption         =   ">"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdPrev 
      Caption         =   "<"
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdLast 
      Caption         =   ">>"
      Height          =   375
      Left            =   4920
      TabIndex        =   7
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdFrist 
      Caption         =   "<<"
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   5160
      Width           =   615
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   5880
      TabIndex        =   5
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtFile 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5295
   End
   Begin VB.ComboBox cboRecords 
      Height          =   300
      Left            =   5640
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   5160
      Width           =   1095
   End
   Begin VB.TextBox txtRecord 
      Height          =   4455
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Width           =   6855
   End
   Begin VB.CommandButton cmdGetRecord 
      Caption         =   "Get Record"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   5160
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_ucmDBF As New cDBF



Private Sub cboRecords_Click()
    GetRecord
End Sub

Private Sub cmdClose_Click()
    If m_ucmDBF.DBFState = stateOpened Then
        cboRecords.Clear
        txtRecord.Text = ""
        m_ucmDBF.DBFClose
        MsgBox "close!"
    Else
    End If
End Sub

Private Sub GetRecord()
    If m_ucmDBF.DBFState = stateOpened Then
        Dim i As Integer
        Dim r() As dbFld
        r = m_ucmDBF.dbfGetRecordEx(cboRecords.Text)
        txtRecord.Text = ""
        For i = LBound(r) To UBound(r)
            txtRecord.Text = txtRecord.Text & r(i).fldName & "=" & r(i).fldValue & vbCrLf
        Next
        'txtRecord.Text = .dbfGetRecord(cboRecords.Text)
        'MsgBox "done!"
    Else
        MsgBox "open file first!"
    End If
End Sub

Private Sub GetRecordEx()
    If m_ucmDBF.DBFState = stateOpened Then
        Dim i As Integer
        Dim r() As dbFld
        r = m_ucmDBF.CurrentRecord
        txtRecord.Text = ""
        For i = LBound(r) To UBound(r)
            txtRecord.Text = txtRecord.Text & r(i).fldName & "=" & r(i).fldValue & vbCrLf
        Next
        cboRecords.ListIndex = m_ucmDBF.CurrentLine - 1
        'txtRecord.Text = .dbfGetRecord(cboRecords.Text)
        If m_ucmDBF.DBFBOF Then
            MsgBox "BOF"
        ElseIf m_ucmDBF.DBFEOF Then
            MsgBox "EOF"
        End If
        
    Else
        MsgBox "open file first!"
    End If
End Sub

Private Sub cmdFrist_Click()
    m_ucmDBF.MoveFirst
    GetRecordEx
End Sub

Private Sub cmdGetRecord_Click()
    GetRecord
End Sub

Private Sub cmdLast_Click()
    m_ucmDBF.MoveLast
    GetRecordEx
End Sub

Private Sub cmdNext_Click()
    m_ucmDBF.MoveNext
    GetRecordEx
End Sub

Private Sub cmdOpen_Click()
    With m_ucmDBF
        If .dbfOpen(txtFile.Text) Then
            Dim o_intItems As Integer
            cboRecords.Visible = False
            For o_intItems = 1 To .RecordCount
                cboRecords.AddItem o_intItems
            Next
            cboRecords.ListIndex = 0
            cboRecords.Visible = True
            m_ucmDBF.CurrentLine = 1
            MsgBox .RecordCount & " records!"
        Else
            MsgBox "file not found!"
        End If
    End With
End Sub

Private Sub cmdPrev_Click()
    m_ucmDBF.MovePrevious
    GetRecordEx
End Sub

Private Sub Form_Load()
    txtFile.Text = App.Path & "\countries.dbf"
End Sub
