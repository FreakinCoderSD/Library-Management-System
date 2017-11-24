VERSION 5.00
Begin VB.Form deleteBook 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Delete A Book"
   ClientHeight    =   3630
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7395
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton search 
      Caption         =   "Search"
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   12
      Top             =   2280
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   11
      Top             =   1800
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   405
      Left            =   1440
      TabIndex        =   10
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Details"
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Left            =   1320
         TabIndex        =   9
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label errorMsg 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Width           =   5340
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Entry Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Call Number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Author Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   1170
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Book Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1080
      End
   End
   Begin VB.TextBox searchCallNo 
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
      Left            =   2160
      MaxLength       =   16
      TabIndex        =   2
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton delBtn 
      Caption         =   "Delete"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
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
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "deleteBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim rs3 As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim sql As String
Dim sql1 As String
Dim sql2 As String

Private Sub CancelButton_Click()
    Unload Me
    Admin.Enabled = True
    Admin.Show
End Sub

Private Sub delBtn_Click()
    If Not rs.EOF And Not rs.BOF Then
        If Not rs2.EOF And rs2.BOF Then
            rs.delete
            Set rs = Nothing
        Else
            rs.delete
             Do Until rs2.EOF
                
                rs2.delete
                rs2.update
                rs2.MoveNext
            Loop
        End If
    MsgBox "Book With It's Copies Are deleted Successfully. ", vbOKOnly, "Deleted Successfully"
    searchCallNo.Text = ""
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    End If
      
End Sub

Private Sub Form_Load()
delBtn.Enabled = False
search.Enabled = False
Admin.Enabled = False

End Sub

Private Sub Form_Unload(cancel As Integer)
Admin.Enabled = True
Admin.Show

End Sub

Private Sub search_Click()
Dim constr As String

    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    Set rs2 = New ADODB.Recordset
    Set rs3 = New ADODB.Recordset
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    con.Open constr
    sql = "SELECT * FROM bookDetails WHERE callNumber = '" & searchCallNo.Text & "'"
    rs.Open sql, con, adOpenDynamic, adLockOptimistic
    If rs.EOF And rs.BOF Then
        MsgBox "Enter a correct call cumber", vbOKOnly, "Wrong call number"
        delBtn.Enabled = False
        rs.Close
    Else
        Text1.Text = rs.Fields(1)
        Text2.Text = rs.Fields(2)
        Text3.Text = rs.Fields(5)
        Text4.Text = rs.Fields(4)
        sql1 = "SELECT * FROM bookShelf WHERE callnumber = '" & rs.Fields(5) & "'"
        rs2.Open sql1, con, adOpenDynamic, adLockOptimistic
        If rs2.EOF And rs2.BOF Then
           errorMsg.Caption = " No copies found of this book "
           delBtn.Enabled = True
        Else
            sql2 = "SELECT * FROM bookShelf WHERE callnumber = '" & rs.Fields(5) & "' AND status = '0'"
            rs3.Open sql2, con, adOpenDynamic, adLockOptimistic
            If Not rs3.EOF And Not rs3.BOF Then
            errorMsg.Caption = "Book having checked out   Copies can't de deleted"
            delBtn.Enabled = False
            Else
            delBtn.Enabled = True
            End If
        End If
    End If
    
        
End Sub

Private Sub searchCallNo_Change()
If Len(searchCallNo.Text) < 16 Then
        search.Enabled = False
ElseIf Len(searchCallNo.Text) = 16 Then
        search.Enabled = True
End If
End Sub

Private Sub searchCallNo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeySpace Then
KeyAscii = 0
End If
End Sub


