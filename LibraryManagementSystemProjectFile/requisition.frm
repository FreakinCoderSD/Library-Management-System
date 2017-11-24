VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form requisitionDetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requisition Details"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7830
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   360
      Top             =   2760
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
      RecordSource    =   "checkoutStatus"
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
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      TabIndex        =   10
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton Detele 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      TabIndex        =   9
      Top             =   2640
      Width           =   1695
   End
   Begin VB.CommandButton checkout 
      Caption         =   "Check Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   2640
      Width           =   1695
   End
   Begin VB.TextBox stdRollBox 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   5655
   End
   Begin VB.TextBox stdNameBox 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   6
      Top             =   1440
      Width           =   5655
   End
   Begin VB.TextBox callNoBox 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   5
      Top             =   840
      Width           =   5655
   End
   Begin VB.TextBox bookNameBox 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label roll 
      AutoSize        =   -1  'True
      Caption         =   "Student Roll"
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
      Left            =   240
      TabIndex        =   3
      Top             =   2040
      Width           =   1320
   End
   Begin VB.Label stdName 
      AutoSize        =   -1  'True
      Caption         =   "Student Name"
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
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label callNo 
      AutoSize        =   -1  'True
      Caption         =   "Call Number"
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
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1290
   End
   Begin VB.Label bookName 
      AutoSize        =   -1  'True
      Caption         =   "Book Name"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1245
   End
End
Attribute VB_Name = "requisitionDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim intResponce As Integer
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim sql As String
Dim constr As String

Private Sub cancel_Click()
Unload Me
Admin.Enabled = True
Admin.Show
End Sub

Private Sub checkout_Click()
intResponce = MsgBox("Are you sure to check out this book", vbYesNo, "Comfirm check out")
If intResponce = vbYes Then
    Adodc1.Refresh
    Adodc1.Recordset.AddNew
    'strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    Adodc1.Recordset.Fields(1) = bookNameBox.Text
    Adodc1.Recordset.Fields(2) = callNoBox.Text
    Adodc1.Recordset.Fields(3) = stdNameBox.Text
    Adodc1.Recordset.Fields(4) = stdRollBox.Text
    Adodc1.Recordset.Fields(5) = Admin.Adodc1.Recordset.Fields(3)
    Adodc1.Recordset.Fields(6) = DateAdd("d", 15, DateValue(Now))
    Adodc1.Recordset.update
    Admin.Adodc1.Recordset.delete
    Unload Me
Else

End If

End Sub

Private Sub Detele_Click()
    intResponce = MsgBox("Are you sure to delete this requisition? ", vbYesNo, "Confirm delete")
    If intResponce = vbYes Then
        
        Set con = New ADODB.Connection
        constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Database.mdb;Persist Security Info=False"
        con.Open constr
        sql = "SELECT * FROM bookshelf WHERE bookUniqueNo = '" & Admin.Adodc1.Recordset.Fields(2) & "'"
        Set rs = New ADODB.Recordset
        rs.Open sql, con, adOpenDynamic, adLockOptimistic
        rs.Fields(4) = 1
        rs.update
        Admin.Adodc1.Recordset.delete
        rs.Close
        con.Close
        Unload Me
    Else
        
    End If
    
End Sub

Private Sub Form_Load()

    Admin.Enabled = False
    
End Sub

Private Sub Form_Unload(cancel As Integer)
    Admin.Enabled = True
    Admin.Show
End Sub

