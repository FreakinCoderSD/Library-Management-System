VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Admin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pro Librarian Administrator "
   ClientHeight    =   9060
   ClientLeft      =   45
   ClientTop       =   645
   ClientWidth     =   16710
   LinkTopic       =   "Form2"
   ScaleHeight     =   9060
   ScaleWidth      =   16710
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   13080
      Top             =   360
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   10800
      Top             =   120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      RecordSource    =   "requisition"
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
   Begin VB.TextBox checkinBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaxLength       =   21
      TabIndex        =   3
      Top             =   120
      Width           =   6375
   End
   Begin VB.CommandButton checkinBtn 
      Caption         =   "Check In"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton checkout 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   14640
      TabIndex        =   1
      Top             =   480
      Width           =   1695
   End
   Begin MSDataGridLib.DataGrid requisitionGrid 
      Bindings        =   "Admin.frx":0000
      Height          =   7665
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   16215
      _ExtentX        =   28601
      _ExtentY        =   13520
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      Enabled         =   -1  'True
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   19
      WrapCellPointer =   -1  'True
      RowDividerStyle =   4
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
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
      Caption         =   "Book Requisitons"
      ColumnCount     =   6
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
         DataField       =   "callNumber"
         Caption         =   "Call Number"
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
         DataField       =   "studentName"
         Caption         =   "Student Name"
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
      BeginProperty Column03 
         DataField       =   "rollNo"
         Caption         =   "Student Roll Number"
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
      BeginProperty Column04 
         DataField       =   "date"
         Caption         =   "Date"
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
      BeginProperty Column05 
         DataField       =   "time"
         Caption         =   "Time"
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
            ColumnWidth     =   3690.142
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   2910.047
         EndProperty
         BeginProperty Column02 
            ColumnWidth     =   2204.788
         EndProperty
         BeginProperty Column03 
            ColumnWidth     =   3209.953
         EndProperty
         BeginProperty Column04 
            ColumnWidth     =   2160
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1725.165
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Call Number"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   1935
   End
   Begin VB.Menu Bk 
      Caption         =   "&Books"
      Index           =   1
      Begin VB.Menu add 
         Caption         =   "Add"
         Index           =   1
         Shortcut        =   ^A
      End
      Begin VB.Menu update 
         Caption         =   "Update Copies"
         Index           =   2
         Shortcut        =   ^U
      End
      Begin VB.Menu delete 
         Caption         =   "Delete"
         Index           =   3
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu account 
      Caption         =   "&Account"
      Index           =   3
      Begin VB.Menu chngPass 
         Caption         =   "Change Password"
         Index           =   1
      End
      Begin VB.Menu logout 
         Caption         =   "Logout"
         Index           =   2
      End
      Begin VB.Menu exit 
         Caption         =   "Logout & Exit"
         Index           =   3
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "Admin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Dim sql As String
Dim constr As String



Private Sub add_Click(Index As Integer)
    Admin.Enabled = False
    addBook.Show
End Sub

Private Sub checkinBox_Change()
    If Len(checkinBox.Text) < 21 Then
        checkinBtn.Enabled = False
    ElseIf Len(checkinBox.Text) = 21 Then
        checkinBtn.Enabled = True
    End If
    
End Sub

Private Sub checkinBtn_Click()
    If Len(checkinBox.Text) = 21 Then
        Set con = New ADODB.Connection
        constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
        con.Open constr
        sql = "SELECT * FROM checkoutStatus WHERE callNumber = '" & checkinBox.Text & "'"
        rs.Open sql, con, adOpenDynamic, adLockOptimistic
        If rs.EOF And rs.BOF Then
            MsgBox "Enter a correct call number", vbSystemModal + vbOKOnly, "Wrong call number"
            rs.Close
        Else
            Admin.Enabled = False
            checkIn.Show
            
        End If
        
    End If
End Sub
Private Sub checkinBox_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
KeyAscii = 0
End If
End Sub

Private Sub checkout_Click()
If Adodc1.Recordset.EOF = True Then
MsgBox "Nothing to check out", vbSystemModal + vbOKOnly, "Nothing to check out"

Else
requisitionDetails.Show
    With requisitionDetails
        .bookNameBox.Text = Adodc1.Recordset.Fields(1)
        .callNoBox.Text = Adodc1.Recordset.Fields(2)
        .stdNameBox.Text = Adodc1.Recordset.Fields(5)
        .stdRollBox.Text = Adodc1.Recordset.Fields(6)
    End With
End If
End Sub



Private Sub chngPass_Click(Index As Integer)
    Admin.Enabled = False
    chngPassword.Show
End Sub

Private Sub delete_Click(Index As Integer)
    deleteBook.Show
    
End Sub

Private Sub exit_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If keyvalue = vbKeyF1 Then
        MsgBox "...."
    End If
End Sub

Private Sub Form_Load()
    checkinBtn.Enabled = False
    Admin.Enabled = False
End Sub


Private Sub Form_Unload(cancel As Integer)
    If frmLogin.loggedin = True Then
        MsgBox " You have successfully logged out. ", vbOKOnly, "Logged out successfully"
    End If
End Sub

Private Sub logout_Click(Index As Integer)
    Admin.Enabled = False
    MsgBox "You have successfully logged out.", vbOKOnly, "Logged out successfully"
    frmLogin.Show
End Sub

Private Sub Timer1_Timer()
    Adodc1.Refresh
End Sub

Private Sub update_Click(Index As Integer)
    updateCopies.Show
    
End Sub
