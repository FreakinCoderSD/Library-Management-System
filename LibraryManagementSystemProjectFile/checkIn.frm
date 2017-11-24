VERSION 5.00
Begin VB.Form checkIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Check In"
   ClientHeight    =   4335
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton reissue 
      Caption         =   "RE-ISSUE"
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox days 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   4800
      TabIndex        =   18
      Top             =   3720
      Width           =   1095
   End
   Begin VB.TextBox fine 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   17
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox checkOut 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   14
      Top             =   3120
      Width           =   3975
   End
   Begin VB.TextBox checkIn 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   13
      Top             =   2520
      Width           =   3975
   End
   Begin VB.TextBox stdRollBox 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   1920
      Width           =   3975
   End
   Begin VB.TextBox stdNameBox 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   11
      Top             =   1320
      Width           =   3975
   End
   Begin VB.TextBox callNoBox 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   720
      Width           =   3975
   End
   Begin VB.TextBox bookName 
      Appearance      =   0  'Flat
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
      Height          =   375
      Left            =   1920
      TabIndex        =   9
      Top             =   120
      Width           =   3975
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "CHECK IN"
      Height          =   375
      Left            =   6000
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label fineDays 
      AutoSize        =   -1  'True
      Caption         =   "Extra Days"
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
      Left            =   3480
      TabIndex        =   16
      Top             =   3720
      Width           =   1155
   End
   Begin VB.Label label 
      AutoSize        =   -1  'True
      Caption         =   "Fine amount"
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
      TabIndex        =   15
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Label checkoutDate 
      AutoSize        =   -1  'True
      Caption         =   "Check Out Date"
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
      TabIndex        =   8
      Top             =   3120
      Width           =   1710
   End
   Begin VB.Label checkinDate 
      AutoSize        =   -1  'True
      Caption         =   "Due Date"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   1035
   End
   Begin VB.Label roll 
      AutoSize        =   -1  'True
      Caption         =   "Student Roll No"
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
      TabIndex        =   6
      Top             =   1920
      Width           =   1680
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
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1245
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
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   1290
   End
End
Attribute VB_Name = "checkIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim sql1 As String
Dim sql2 As String
Dim diff As Long
Dim fineAmount As Integer

Private Sub CancelButton_Click()
Unload Me
Admin.Enabled = True
Admin.Show

End Sub

Private Sub Form_Load()
Admin.Enabled = False
Call fine_calculate
bookName.Text = Admin.rs.Fields(1)
callNoBox.Text = Admin.rs.Fields(2)
stdNameBox.Text = Admin.rs.Fields(3)
stdRollBox.Text = Admin.rs.Fields(4)
checkIn.Text = Admin.rs.Fields(6)
checkOut.Text = Admin.rs.Fields(5)
fine.Text = fineAmount
days.Text = diff

End Sub

Public Function fine_calculate()

diff = DateDiff("d", Admin.rs.Fields(6), DateValue(Now))
If diff <= 0 Then
    fineAmount = 0#
    diff = 0
Else
    fineAmount = 2# * diff
    
End If

End Function

Private Sub Form_Unload(cancel As Integer)
    Admin.Enabled = True
    Admin.con.Close
    Admin.Show
    Admin.checkinBox.Text = ""
    Admin.checkinBox.SetFocus
End Sub

Private Sub OKButton_Click()
Dim check As Integer
check = MsgBox("Are You sure to check in this book", vbYesNo, "Confirm check in")
If check = vbYes Then
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    rs.Open "SELECT * FROM bookshelf WHERE bookUniqueNo ='" & callNoBox.Text & "'", con, adOpenDynamic, adLockOptimistic
    rs.Fields(4) = "1"
    rs.update
    Admin.rs.delete
    rs.Close
    Admin.rs.Close
    Unload Me
Else
End If
End Sub

Private Sub reissue_Click()
Dim check As Integer
check = MsgBox("Are you sure To re-issue this book for next 15 days", vbYesNo, "Confirm re-issue")
If check = vbYes Then
    Admin.rs.Fields(5) = DateValue(Now)
    Admin.rs.Fields(6) = DateAdd("d", 15, DateValue(Now))
    Admin.rs.update
    Admin.rs.Close
    Unload Me
Else

End If
End Sub
