VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDBF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Option Base 1

'Portions adopted from Erik Olsen's PB3-DBF.BAS
'Portions Copyright (c) 1992, Ethan Winer
'Portions Copyright (c) 2000, Steven A. Tanimura

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type fldHeaderRecord
    ver As Byte 'fld version
    yy As Byte 'year
    mm As Byte 'month
    dd As Byte 'day of last update
    recs As Long 'number of records in this database
    offset As Integer 'length of header
    rLen As Integer 'length of record
    dummy As String * 20 'reserved for future use
End Type

Private Type fldFldRecStruc
    fldName As String * 11 'name of the fld in ASCII
    fldType As String * 1 'Type CNLM or D
    fda As Long 'fld data address - we don't need this
    fLen As Byte 'Length, we'll need this!
    decC As Byte 'number of decimals in numeric fld
    blank As String * 14 'reserved for future use
End Type

Private Type dbFldArray
    fldName As String * 11
    fldType As String * 1
    fldLength As Byte
    fldOffset As Integer
    fldDecimals As Byte
End Type

'Private memo As String * 512

Public Enum dbFldType
    Character 'C
    [Currency] 'Y
    Numeric 'N
    Float 'F
    Date 'D
    [DateTime] 'T
    [Double] 'D
    [Binary] 'B Binary in dBASE
    [Integer] 'I
    Logical 'L
    Memo 'M
    General 'G
    Picture 'P
    NullFlag '0
End Enum

Public Type dbFld
    fldName As String
    fldValue As Variant
    fldType As dbFldType
End Type

Public Enum dbState
    stateOpened
    stateNotReady
End Enum

Private m_udeState As dbState
Private m_udtHeader As fldHeaderRecord 'header
Private m_udtField As fldFldRecStruc 'field
Private m_udtFieldArray(1 To 25) As dbFldArray 'data field array
Private m_intFileNum As Long 'file handle
Private m_lngRecords As Long ' total number of records in the current database
Private m_lngFields As Integer 'total number of fields in the current database
Private m_lngCurrentLine As Integer
Private m_udtCurrentFields() As dbFld
Private m_blnEOF As Boolean
Private m_blnBOF As Boolean



Public Property Get DBFBOF() As Boolean
    DBFBOF = m_blnBOF
End Property

Public Property Get DBFEOF() As Boolean
    DBFEOF = m_blnEOF
End Property

Public Property Get DBFState() As dbState
    DBFState = m_udeState
End Property

Public Function DBFOpen(ByVal strFileName As String) As Boolean
    
    If Dir$(strFileName) <> "" And m_udeState = stateNotReady Then
        Dim o_intField As Integer
        Dim o_strRet As String

        m_intFileNum = FreeFile
    
        Open strFileName For Binary As #m_intFileNum
    
        If LOF(m_intFileNum) = 0 Then
            Close #m_intFileNum
        End If
    
        Get #m_intFileNum, 1, m_udtHeader
    
        m_lngRecords = m_udtHeader.recs
        m_lngFields = (m_udtHeader.offset - 32) \ 32
        
        If m_lngFields < 1 Then
            Close #m_intFileNum
        End If
    
        'Load the m_udtField definition header
        m_udtFieldArray(1).fldOffset = 2
    
        For o_intField = 1 To m_lngFields
            Get #m_intFileNum, , m_udtField
            
            With m_udtField
                o_strRet = TrimEx(.fldName)
                
                If o_strRet <> "" Then
                    m_udtFieldArray(o_intField).fldName = .fldName
                    m_udtFieldArray(o_intField).fldType = .fldType
                    
                    m_udtFieldArray(o_intField).fldLength = .fLen
                    
                    If o_intField < m_lngFields Then
                       m_udtFieldArray(o_intField + 1).fldOffset = m_udtFieldArray(o_intField).fldOffset + .fLen
                    End If
                    
                    m_udtFieldArray(o_intField).fldDecimals = .decC
                Else
                    m_lngFields = o_intField - 1
                    Exit For
                End If
            End With
        Next
        
        m_udeState = stateOpened
        
        DBFOpen = True
    Else
        DBFOpen = False
    End If
  
End Function

Public Sub DBFClose()
    If m_udeState = stateOpened Then
        m_udeState = stateNotReady
        Close #m_intFileNum
    Else
    End If
End Sub

Public Property Get CurrentLine() As Long
    CurrentLine = m_lngCurrentLine
End Property

Public Property Let CurrentLine(ByVal lngCurrentLine As Long)
    m_lngCurrentLine = lngCurrentLine
    
    MoveTo m_lngCurrentLine
End Property

Public Function MoveTo(ByVal lngRecNum As Long) As Boolean
    If lngRecNum > 0 And lngRecNum <= m_lngRecords And m_udeState = stateOpened Then
        m_lngCurrentLine = lngRecNum
        m_udtCurrentFields() = DBFGetRecordEx(m_lngCurrentLine)
        MoveTo = True
    Else
        MoveTo = False
    End If
End Function

Public Function MoveFirst() As Boolean
    Dim o_blnRet As Boolean
    If m_udeState = stateOpened Then
        If m_lngRecords > 0 Then
            m_lngCurrentLine = 1
            SetState False, False
            o_blnRet = MoveTo(m_lngCurrentLine)
        Else
            SetState True, True
            o_blnRet = True
        End If
    Else
        o_blnRet = False
    End If
    MoveFirst = o_blnRet
End Function

Public Function MoveLast() As Boolean
    Dim o_blnRet As Boolean
    If m_udeState = stateOpened Then
        If m_lngRecords > 0 Then
            m_lngCurrentLine = m_lngRecords
            SetState False, False
            o_blnRet = MoveTo(m_lngCurrentLine)
        Else
            SetState True, True
            o_blnRet = True
        End If
    Else
        o_blnRet = False
    End If
    MoveLast = o_blnRet
End Function

Public Function MovePrevious() As Boolean
    Dim o_blnRet As Boolean
    If m_udeState = stateOpened Then
        If m_lngRecords > 0 Then
            If m_lngCurrentLine > 1 Then
                SetState False, False
                m_lngCurrentLine = m_lngCurrentLine - 1
                o_blnRet = MoveTo(m_lngCurrentLine)
            Else
                SetState True, False
                o_blnRet = True
            End If
        Else
            SetState True, True
            o_blnRet = True
        End If
    Else
        o_blnRet = False
    End If
    MovePrevious = o_blnRet
End Function

Public Function MoveNext() As Boolean
    Dim o_blnRet As Boolean
    If m_udeState = stateOpened Then
        If m_lngRecords > 0 Then
            If m_lngCurrentLine < m_lngRecords Then
                SetState False, False
                m_lngCurrentLine = m_lngCurrentLine + 1
                o_blnRet = MoveTo(m_lngCurrentLine)
            Else
                SetState False, True
                o_blnRet = True
            End If
        Else
            SetState True, True
            o_blnRet = True
        End If
    Else
        o_blnRet = False
    End If
    MoveNext = o_blnRet
End Function

Private Sub SetState(ByVal blnBOF As Boolean, ByVal blnEOF As Boolean)
    m_blnBOF = blnBOF
    m_blnEOF = blnEOF
End Sub

Public Function CurrentRecord() As dbFld()
    CurrentRecord = m_udtCurrentFields()
End Function

Public Property Get RecordCount() As Long
    RecordCount = m_lngRecords
End Property

Public Property Get FieldCount() As Long
    FieldCount = m_lngFields
End Property

Public Function DBFGetRecord(ByVal lngRecNum As Long) As String
    If m_udeState = stateOpened Then
        Dim o_lngOffset As Long
        Dim o_strRet As String

        o_strRet = String$(m_udtHeader.rLen, " ")
        o_lngOffset = RecordOffset(lngRecNum)
    
        Get #m_intFileNum, o_lngOffset, o_strRet
        
        DBFGetRecord = o_strRet
    Else
        DBFGetRecord = vbNull
    End If
End Function

Public Function DBFGetRecordEx(ByVal lngRecNum As Long) As dbFld()
    If m_udeState = stateOpened Then
        Dim o_lngOffset As Long
        Dim o_byteContent() As Byte
        Dim o_byteField() As Byte
        Dim o_lngItems As Long
        Dim o_strRet As String
        Dim o_udeType As dbFldType
        Dim o_udtFields() As dbFld
        Dim o_vntValue As Variant
        ReDim o_udtFields(1 To m_lngFields)
        
        o_lngOffset = RecordOffset(lngRecNum)
        
        ReDim o_byteContent(0 To m_udtHeader.rLen - 1)
            
        'Seek #m_intFileNum, o_lngOffset
        
        Get #m_intFileNum, o_lngOffset, o_byteContent
        
        For o_lngItems = 1 To m_lngFields
            With m_udtFieldArray(o_lngItems)
                o_udtFields(o_lngItems).fldName = TrimEx(.fldName)
                
                ReDim o_byteField(0 To .fldLength - 1)
                
                CopyMemory o_byteField(0), o_byteContent(.fldOffset - 1), .fldLength
                
                o_strRet = TrimEx(StrConv(o_byteField, vbUnicode))
                
                Select Case UCase(.fldType)
                    Case "C"
                        o_udeType = Character
                        o_vntValue = CStr(o_strRet)
                    Case "Y"
                        o_udeType = [Currency]
                        o_vntValue = CCur(o_strRet)
                    Case "N"
                        o_udeType = Numeric
                        o_vntValue = Val(o_strRet)
                    Case "F"
                        o_udeType = [Float]
                        o_vntValue = Val(o_strRet)
                    Case "D"
                        o_udeType = [DateTime]
                        o_vntValue = CDate(o_strRet)
                    Case "B"
                        o_udeType = Binary
                        o_vntValue = CStr(o_strRet) 'to do
                    Case "I"
                        o_udeType = [Integer]
                        o_vntValue = CInt(o_strRet)
                    Case "L"
                        o_udeType = Logical
                        o_vntValue = CBool(o_strRet)
                    Case "M"
                        o_udeType = Memo
                        o_vntValue = o_strRet 'to do
                    Case "P"
                        o_udeType = Picture
                        o_vntValue = o_strRet 'to do
                    Case "0"
                        o_udeType = NullFlag
                        o_vntValue = vbNull
                    Case Else 'G
                        o_udeType = General
                        o_vntValue = o_strRet
                End Select
                
                o_udtFields(o_lngItems).fldValue = o_vntValue
            End With
        Next
    
        DBFGetRecordEx = o_udtFields
    Else
    End If
End Function

Public Function DBFSetField(ByVal strRecord As String, ByVal strFieldText As String, ByVal strFieldNum As Integer) As Boolean
    If m_udeState = stateOpened Then
    '    Dim temp As String
    '
    '    temp = Space$(m_udtFieldArray(strFieldNum).fldLength)
    '    LSet temp = strFieldText
        Mid$(strRecord, m_udtFieldArray(strFieldNum).fldOffset, m_udtFieldArray(strFieldNum).fldLength) = strFieldText
        DBFSetField = True
    Else
        DBFSetField = False
    End If
End Function

Public Function DBFPutRecord(ByVal lngRecNum As Long, ByVal strRecord As String) As Boolean
    If m_udeState = stateOpened Then
        Dim o_lngOffset As Long
    
        o_lngOffset = RecordOffset(lngRecNum)
    
        Put #m_intFileNum, o_lngOffset, strRecord
        
        DBFPutRecord = True
    Else
        DBFPutRecord = False
    End If
End Function

Private Function RecordOffset(ByVal lngRecNum As Long) As Long
    RecordOffset = ((lngRecNum - 1) * m_udtHeader.rLen) + m_udtHeader.offset + 1
End Function

Public Sub PackHeaderDate()
    Dim today As String
'    Dim yr As Byte
'    Dim mo As Byte
'    Dim dy As Byte

    today = CStr(Date)
    m_udtHeader.yy = CByte(Val(Right$(today, 2)))
    m_udtHeader.mm = CByte(Val(Left$(today, 2)))
    m_udtHeader.dd = CByte(Val(Mid$(today, 4, 2)))
End Sub

Public Function GetHeaderDate() As String
    Dim yr As String * 4
    Dim mo As String * 2
    Dim da As String * 2

    yr = Str$(2000 + m_udtHeader.yy)
    mo = Right$("0" + Str$(m_udtHeader.mm), 2)
    da = Right$("0" + Str$(m_udtHeader.dd), 2)

    GetHeaderDate = yr & mo & da
End Function

Public Function DBFDelRecord(ByVal strRecord As String) As Boolean
    If m_udeState = stateOpened Then
        Mid$(strRecord, 1) = "*"
        DBFDelRecord = True
    Else
        DBFDelRecord = False
    End If
End Function

Public Function DBFUpdateHdr(ByVal NbrRecs As Long) As Boolean
    If m_udeState = stateOpened Then
        Put #m_intFileNum, 5, NbrRecs
        DBFUpdateHdr = True
    Else
        DBFUpdateHdr = False
    End If
End Function

Private Sub Class_Initialize()
    m_udeState = stateNotReady
End Sub

Private Sub Class_Terminate()
    DBFClose
End Sub

Private Function TrimEx(ByVal strText As String) As String
    If strText <> "" Then
        Dim o_lngRet As Long
        o_lngRet = InStr(1, strText, Chr(0))
        If o_lngRet > 0 Then
            TrimEx = Trim(Left(strText, o_lngRet - 1))
            o_lngRet = InStr(1, strText, Chr(13))
            If o_lngRet > 0 Then
                TrimEx = Trim(Left(strText, o_lngRet - 1))
            Else
            End If
        Else
            TrimEx = Trim(strText)
        End If
    Else
    End If
End Function

