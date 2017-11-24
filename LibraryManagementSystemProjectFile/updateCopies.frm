VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form updateCopies 
   Caption         =   "Update Copies"
   ClientHeight    =   5490
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   9105
   LinkTopic       =   "Form2"
   ScaleHeight     =   5490
   ScaleWidth      =   9105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton addbtn 
      Caption         =   "Add Copy"
      Height          =   375
      Left            =   7560
      TabIndex        =   5
      Top             =   1440
      Width           =   1335
   End
   Begin VB.CommandButton deleteBtn 
      Caption         =   "Delete Copy"
      Height          =   375
      Left            =   7560
      TabIndex        =   4
      Top             =   2160
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7560
      Top             =   600
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SELECT * FROM bookshelf"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "updateCopies.frx":0000
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7858
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   3
      BeginProperty Column00 
         DataField       =   "bookName"
         Caption         =   "Book Name"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "bookUniqueNo"
         Caption         =   "Call Nummber"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "status"
         Caption         =   "status"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
            ColumnWidth     =   2940.095
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2835.213
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton check 
      Caption         =   "Check"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox callNo 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3000
      MaxLength       =   16
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter The Call Number"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2565
   End
End
Attribute VB_Name = "updateCopies"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim copyCount As Long
Private Sub addbtn_Click()
    Dim bookName As String
    Dim callNo As String
    Dim uniqueNo As String
    bookName = Adodc1.Recordset.Fields(1)
    callNo = Adodc1.Recordset.Fields(2)
    
    Adodc1.Recordset.MoveLast
    copyCount = Right$(Adodc1.Recordset.Fields(3), 3)
    uniqueNo = CStr(Format(copyCount + 1, "000"))
    Adodc1.Recordset.AddNew
      Adodc1.Recordset.Fields(1) = bookName
      Adodc1.Recordset.Fields(2) = callNo
      Adodc1.Recordset.Fields(3) = callNo + "_c" + uniqueNo
      Adodc1.Recordset.Fields(4) = "1"
    
End Sub
Private Sub callNo_Change()
If Len(callNo.Text) < 16 Then
        check.Enabled = False
    ElseIf Len(callNo.Text) = 16 Then
        check.Enabled = True
End If
End Sub

Private Sub callNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
KeyAscii = 0
End If
End Sub

Private Sub check_Click()
    If Len(callNo.Text) = 16 Then
        Adodc1.RecordSource = "SELECT * FROM bookshelf WHERE callNumber = '" & callNo.Text & "'"
        Adodc1.Refresh
    Else
        MsgBox " not a valid call Number"
    End If
    If Adodc1.Recordset.EOF And Adodc1.Recordset.BOF Then
        DataGrid1.Enabled = False
        MsgBox "Book Not Found "
    Else
        addbtn.Enabled = True
        deleteBtn.Enabled = True
        DataGrid1.Enabled = True
    End If
End Sub

Private Sub deleteBtn_Click()
    Dim i As Integer
    If Adodc1.Recordset.Fields(4) = 1 Then
        i = MsgBox("Are you sure to delete this copy", vbYesNo, "Confirm delete")
        If i = vbYes Then
          
        Adodc1.Recordset.delete
        
        Else
        
        End If
    Else
        MsgBox ("Checked out book can't be deleted")
    End If
End Sub

Private Sub Form_Load()
    Admin.Enabled = False
    addbtn.Enabled = False
    deleteBtn.Enabled = False
   DataGrid1.ClearFields
    
End Sub

Private Sub Form_Unload(cancel As Integer)
Admin.Enabled = True
Admin.Show
End Sub
