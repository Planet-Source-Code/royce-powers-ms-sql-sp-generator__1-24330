VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Generate Stored Procedures from Tables"
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11445
   Icon            =   "sp_Gen.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7545
   ScaleWidth      =   11445
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkApply 
      Caption         =   "Apply to Database"
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CheckBox chkAll 
      Caption         =   "Check/Uncheck All"
      Height          =   375
      Left            =   3360
      TabIndex        =   13
      Top             =   240
      Width           =   2895
   End
   Begin VB.ListBox lstTables 
      Height          =   5910
      Left            =   3360
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   720
      Width           =   4095
   End
   Begin VB.TextBox txtSP 
      Height          =   6375
      Left            =   7560
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   240
      Width           =   3855
   End
   Begin VB.CommandButton cmdButtons 
      Caption         =   "Generate"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   10
      Top             =   240
      Width           =   975
   End
   Begin VB.Frame fraSQL 
      Caption         =   " SQL Server "
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3135
      Begin VB.CommandButton cmdButtons 
         Caption         =   "Connect"
         Height          =   495
         Index           =   0
         Left            =   1080
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox txtInputs 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1080
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1440
         Width           =   1575
      End
      Begin VB.TextBox txtInputs 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox txtInputs 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   6
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox txtInputs 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Password:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "User ID:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Database:"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Server:"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Copyright 2001, Royce D. Powers"
      Height          =   375
      Left            =   120
      TabIndex        =   15
      Top             =   7080
      Width           =   7815
   End
   Begin VB.Menu mnuHelpAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Cn As New ADODB.Connection
Dim Rs As New ADODB.Recordset
Dim Cm As New ADODB.Command

Private Sub chkAll_Click()
    Dim intCount As Integer
    For intCount = 0 To lstTables.ListCount - 1
        lstTables.Selected(intCount) = chkAll.Value
    Next intCount
End Sub

Private Sub cmdButtons_Click(Index As Integer)
    
    Select Case Index
        Case 0 'Connect
            Call ConnectSQL
        Case 1 'Generate
            Call Generate
    End Select
End Sub

Private Function ConnectSQL() As String
On Error GoTo ErrorHandler

    Dim intCount As Integer
    Dim strConnectionString As String
    
    strConnectionString = "Driver={SQL Server};"
    strConnectionString = strConnectionString & "Server=" & txtInputs(0) & ";"
    strConnectionString = strConnectionString & "Uid=" & txtInputs(2) & ";"
    strConnectionString = strConnectionString & "Pwd=" & txtInputs(3) & ";"
    strConnectionString = strConnectionString & "Database=" & txtInputs(1)
   
    'Open the connection
    With Cn
        .Open strConnectionString
        .CursorLocation = adUseClient
    End With

    ' Get the schema
    Set Rs = Cn.OpenSchema(adSchemaTables, Array(Empty, Empty, Empty, "TABLE"))
    
    ' Load the tables into the list box
    While Not Rs.EOF
        lstTables.AddItem Rs(2)
        Rs.MoveNext
    Wend
    Cm.Prepared = True
    Cm.ActiveConnection = Cn
    Exit Function
ErrorHandler:
    MsgBox Err.Number & vbCrLf & Err.Description, vbCritical, "Database Connection Error"
End Function

Private Sub Generate()
    Dim intCount As Integer
    Dim strSQL As String
    
    Me.MousePointer = vbHourglass
    ' Get the table names
    For intCount = 0 To lstTables.ListCount - 1
        If lstTables.Selected(intCount) Then
            ' Get the primary keys
            txtSP.Text = ""
            strSQL = strSQL & CreateSelectProc(lstTables.List(intCount))
            strSQL = strSQL & CreateInsertProc(lstTables.List(intCount))
            strSQL = strSQL & CreateUpdateProc(lstTables.List(intCount))
            strSQL = strSQL & CreateDeleteProc(lstTables.List(intCount))
        End If
    Next intCount
    txtSP.Text = strSQL
    Me.MousePointer = vbDefault
End Sub

Private Function CreateSelectProc(strTable As String) As String
    Dim strSQLDrop As String
    Dim strSQLProc As String
    strSQLDrop = "USE " & txtInputs(1).Text & vbCrLf
    strSQLDrop = strSQLDrop & "IF EXISTS (SELECT name FROM sysobjects" & vbCrLf
    strSQLDrop = strSQLDrop & "WHERE name = 'sp_Select_" & strTable & "' AND type = 'P')"
    strSQLDrop = strSQLDrop & "DROP PROCEDURE [sp_Select_" & strTable & "]" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = strSQLDrop
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "Go" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    strSQLProc = strSQLProc & "CREATE PROCEDURE [sp_Select_" & strTable & "] AS " & vbCrLf
    strSQLProc = strSQLProc & "SET NOCOUNT ON" & vbCrLf
    strSQLProc = strSQLProc & "SELECT * FROM [" & strTable & "]" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = Right(strSQLProc, Len(strSQLProc) - 3)
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "GO" & vbCrLf & vbCrLf
    CreateSelectProc = strSQLDrop & strSQLProc
End Function

Private Function CreateInsertProc(strTable As String) As String
    Dim strSQLDrop As String
    Dim strSQLProc As String
    Dim intCount As Integer
    Dim RsIndexes As New ADODB.Recordset
    
    
    ' Write DROP and header    ...
    strSQLDrop = "USE " & txtInputs(1) & vbCrLf
    strSQLDrop = strSQLDrop & "IF EXISTS (SELECT name FROM sysobjects" & vbCrLf
    strSQLDrop = strSQLDrop & "WHERE name = 'sp_Insert_" & strTable & "' AND type = 'P')"
    strSQLDrop = strSQLDrop & "DROP PROCEDURE [sp_Insert_" & strTable & "]" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = strSQLDrop
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "Go" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    strSQLProc = strSQLProc & "CREATE PROCEDURE [sp_Insert_" & strTable & "]" & vbCrLf
    
    'Add Columns
    Set Rs = Cn.Execute("sp_Columns '" & strTable & "'")
        While Not Rs.EOF
            strSQLProc = strSQLProc & "@" & Rs(3) & " "
            Select Case Rs(4)
                Case 4 ' Identity
                    strSQLProc = strSQLProc & "int"
                Case Is < 0
                    strSQLProc = strSQLProc & Rs(5)
                    If Rs(4) <> -10 And Rs(4) <> -4 And Rs(4) <> -7 Then strSQLProc = strSQLProc & " " & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
                Case 12
                    strSQLProc = strSQLProc & " " & Rs(5) & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
                Case Else
                    strSQLProc = strSQLProc & Rs(5)
            End Select
            If Rs(10) = 1 Then strSQLProc = strSQLProc & " = Null"
            
            If intCount < Rs.RecordCount - 1 Then
                strSQLProc = strSQLProc & ","
                intCount = intCount + 1
            End If
            strSQLProc = strSQLProc & vbCrLf
            Rs.MoveNext
        Wend
   
    strSQLProc = strSQLProc & "AS " & vbCrLf
    strSQLProc = strSQLProc & "SET NOCOUNT ON" & vbCrLf
    strSQLProc = strSQLProc & "Begin Transaction" & vbCrLf
    strSQLProc = strSQLProc & "INSERT [" & strTable & "]" & vbCrLf & " ("
    
    Set Rs = Cn.Execute("sp_Columns '" & strTable & "'")
    intCount = 0
        While Not Rs.EOF
            strSQLProc = strSQLProc & Rs(3)
            If intCount < Rs.RecordCount - 1 Then strSQLProc = strSQLProc & "," & vbCrLf
            strSQLProc = strSQLProc & " "
            Rs.MoveNext
            intCount = intCount + 1
        Wend
    strSQLProc = strSQLProc & ")" & vbCrLf
    strSQLProc = strSQLProc & "VALUES " & vbCrLf & "("
    
    Set Rs = Cn.Execute("sp_Columns '" & strTable & "'")
    intCount = 0
        While Not Rs.EOF
            strSQLProc = strSQLProc & "@" & Rs(3)
            If intCount < Rs.RecordCount - 1 Then
                strSQLProc = strSQLProc & ", " & vbCrLf
            End If
            intCount = intCount + 1
            Rs.MoveNext
        Wend
    strSQLProc = strSQLProc & ") " & vbCrLf
    strSQLProc = strSQLProc & "COMMIT" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = Right(strSQLProc, Len(strSQLProc) - 3)
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "GO" & vbCrLf & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    CreateInsertProc = strSQLDrop & strSQLProc
End Function

Private Function CreateUpdateProc(strTable As String) As String
    Dim strSQLDrop As String
    Dim strSQLProc As String
    Dim intCount As Integer
    Dim RsIndexes As New ADODB.Recordset
    
    
    ' Write DROP and header    ...
    strSQLDrop = "USE " & txtInputs(1) & vbCrLf
    strSQLDrop = strSQLDrop & "IF EXISTS (SELECT name FROM sysobjects" & vbCrLf
    strSQLDrop = strSQLDrop & "WHERE name = 'sp_Update_" & strTable & "' AND type = 'P')"
    strSQLDrop = strSQLDrop & "DROP PROCEDURE [sp_Update_" & strTable & "]" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = strSQLDrop
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "Go" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    strSQLProc = strSQLProc & "CREATE PROCEDURE [sp_Update_" & strTable & "]" & vbCrLf

    ' Get the primary key
    Set Rs = Cn.Execute("sp_Columns '" & strTable & "'")
    While Not Rs.EOF
        strSQLProc = strSQLProc & "@" & Rs(3) & " "
        Select Case Rs(4)
            Case 4 ' Identity
                strSQLProc = strSQLProc & "int"
            Case Is < 0
                strSQLProc = strSQLProc & Rs(5)
                If Rs(4) <> -10 And Rs(4) <> -4 And Rs(4) <> -7 Then strSQLProc = strSQLProc & " " & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
            Case 12
                strSQLProc = strSQLProc & " " & Rs(5) & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
            Case Else
                strSQLProc = strSQLProc & Rs(5)
        End Select
        If Rs(10) = 1 Then strSQLProc = strSQLProc & " = Null"
        
        If intCount < Rs.RecordCount - 1 Then
            strSQLProc = strSQLProc & ","
            intCount = intCount + 1
        End If
        strSQLProc = strSQLProc & vbCrLf
        Rs.MoveNext
    Wend
    
    strSQLProc = strSQLProc & "AS " & vbCrLf
    strSQLProc = strSQLProc & "SET NOCOUNT ON" & vbCrLf
    strSQLProc = strSQLProc & "Begin Transaction" & vbCrLf
    strSQLProc = strSQLProc & "UPDATE [" & strTable & "]" & vbCrLf
    strSQLProc = strSQLProc & "SET " & vbCrLf
    
    Rs.MoveFirst
    intCount = 0
    While Not Rs.EOF
        strSQLProc = strSQLProc & Rs(3) & " = " & "@" & Rs(3)
        Rs.MoveNext
        If intCount < Rs.RecordCount - 1 Then strSQLProc = strSQLProc & "," & vbCrLf
        intCount = intCount + 1
    Wend
    
    strSQLProc = strSQLProc & vbCrLf

    strSQLProc = strSQLProc & "WHERE "
    Set Rs = Cn.Execute("sp_PKeys '" & strTable & "'")
    intCount = 0
    While Not Rs.EOF
        strSQLProc = strSQLProc & Rs(3) & " = @" & Rs(3)
        If intCount < Rs.RecordCount - 1 Then strSQLProc = strSQLProc & " AND "
        intCount = intCount + 1
        Rs.MoveNext
    Wend
    strSQLProc = strSQLProc & vbCrLf
    strSQLProc = strSQLProc & "COMMIT" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = Right(strSQLProc, Len(strSQLProc) - 3)
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "GO" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    
    CreateUpdateProc = strSQLDrop & strSQLProc
End Function

Private Function CreateDeleteProc(strTable As String) As String
    Dim strSQLDrop As String
    Dim strSQLProc As String
    Dim intCount As Integer
    Dim RsIndexes As New ADODB.Recordset
    
    
    ' Write DROP and header    ...
    strSQLDrop = "USE " & txtInputs(1) & vbCrLf
    strSQLDrop = strSQLDrop & "IF EXISTS (SELECT name FROM sysobjects" & vbCrLf
    strSQLDrop = strSQLDrop & "WHERE name = 'sp_Delete_" & strTable & "' AND type = 'P')"
    strSQLDrop = strSQLDrop & "DROP PROCEDURE [sp_Delete_" & strTable & "]" & vbCrLf
    If chkApply.Value = 1 Then
        Cm.CommandText = strSQLDrop
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "Go" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    strSQLProc = strSQLProc & "CREATE PROCEDURE [sp_Delete_" & strTable & "]" & vbCrLf
    
    Set Rs = Cn.Execute("sp_Columns '" & strTable & "'")
    Set RsIndexes = Cn.Execute("sp_PKeys '" & strTable & "'")
    
    intCount = 0
    While Not Rs.EOF
        RsIndexes.MoveFirst
        While Not RsIndexes.EOF
            If Rs(3) = RsIndexes(3) Then
                strSQLProc = strSQLProc & "@" & Rs(3) & " "
                Select Case Rs(4)
                    Case 4 ' Identity
                        strSQLProc = strSQLProc & "int"
                    Case Is < 0
                        strSQLProc = strSQLProc & Rs(5)
                        If Rs(4) <> -10 And Rs(4) <> -4 And Rs(4) <> -7 Then strSQLProc = strSQLProc & " " & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
                    Case 12
                        strSQLProc = strSQLProc & " " & Rs(5) & IIf(Rs(7) > 255, "(" & 255 & ")", "(" & Rs(7) & ")")
                    Case Else
                        strSQLProc = strSQLProc & Rs(5)
                End Select
                If Rs(10) = 1 Then strSQLProc = strSQLProc & " = Null"
                
                If intCount < RsIndexes.RecordCount - 1 Then
                    strSQLProc = strSQLProc & ","
                    intCount = intCount + 1
                End If
                strSQLProc = strSQLProc & vbCrLf
            End If
            RsIndexes.MoveNext
        Wend
        Rs.MoveNext
    Wend
        
    strSQLProc = strSQLProc & "AS" & vbCrLf
    strSQLProc = strSQLProc & "DELETE FROM [" & strTable & "]" & vbCrLf
    strSQLProc = strSQLProc & "WHERE "
    
    RsIndexes.MoveFirst
    intCount = 0
    While Not RsIndexes.EOF
        strSQLProc = strSQLProc & RsIndexes(3) & " = @" & RsIndexes(3)
        If intCount < RsIndexes.RecordCount - 1 Then strSQLProc = strSQLProc & " AND "
        intCount = intCount + 1
        RsIndexes.MoveNext
    Wend
    
    strSQLProc = strSQLProc & vbCrLf
    strSQLProc = strSQLProc & "COMMIT" & vbCrLf
        If chkApply.Value = 1 Then
        Cm.CommandText = Right(strSQLProc, Len(strSQLProc) - 3)
        Cm.Execute
    End If
    strSQLProc = strSQLProc & "GO" & vbCrLf
    strSQLProc = strSQLProc & vbCrLf & vbCrLf
    
    CreateDeleteProc = strSQLDrop & strSQLProc
End Function

Private Sub mnuHelpAbout_Click()
    frmAbout.Show vbModal, Me
End Sub
