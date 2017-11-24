VERSION 5.00
Begin VB.Form frmSignUp 
   Caption         =   "frmSignUp"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   ScaleHeight     =   3675
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton reset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   8
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton ok 
      Caption         =   "Ok"
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   2160
      Width           =   1335
   End
   Begin VB.TextBox password1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   1560
      Width           =   3375
   End
   Begin VB.TextBox password 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   960
      Width           =   3375
   End
   Begin VB.TextBox userName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "NOTE: Spaces are not available in both username and password."
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
      Height          =   240
      Left            =   -840
      TabIndex        =   11
      Top             =   3360
      Width           =   7755
      WordWrap        =   -1  'True
   End
   Begin VB.Label msg 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   240
      TabIndex        =   10
      Top             =   2160
      Width           =   45
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1170
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "UserName"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "You Seem To Be a New User Of This Software. Please Create an Username And Password  For Future Login."
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
      Height          =   480
      Left            =   105
      TabIndex        =   0
      Top             =   2760
      Width           =   5775
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmSignUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim str As String
Private Sub cancel_Click()
    Unload Me
End Sub

Private Sub ok_Click()
    If userName.Text = "" Then
        msg.Caption = ""
        MsgBox "Please Enter a UserName. ", vbOKOnly, "Enter UserName"
        userName.SetFocus
        
    ElseIf password.Text = "" Then
        msg.Caption = ""
        MsgBox "Please Enter a New PassWord. ", vbOKOnly, "Enter Password"
        password.SetFocus
    ElseIf password1.Text = "" Then
        msg.Caption = ""
        MsgBox "Please Confirm Your password.", vbOKOnly, "Confirm Password"
        password1.SetFocus
    ElseIf password.Text <> password1.Text Then
    MsgBox "Password & Confirm Password doesn't Match", vbOKOnly, "Try Again"
        password.Text = ""
        password1.Text = ""
    Else
        msg.Caption = ""
        Set con = New ADODB.Connection
        str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=TRUE"
        con.Open str
        Set rs = New ADODB.Recordset
        str = "SELECT * FROM admin"
        rs.Open str, con, adOpenDynamic, adLockOptimistic
        rs.AddNew
        rs.Fields(1) = userName.Text
        rs.Fields(2) = password.Text
        rs.update
        Admin.Enabled = True
        MsgBox "UserName & Password Saved For future Login.", vbOKOnly, "UserName and Password saved."
        Unload Me
        Admin.Enabled = True
        Admin.Show
        frmLogin.loggedin = True
        
        
     End If
End Sub
Private Sub Form_Unload(cancel As Integer)
    frmLogin.loggedin = False
    If rs.State = 1 Then
    rs.Close
    con.Close
    End If
    frmLogin.loggedin = False
    
    Unload Admin
 
End Sub


Private Sub reset_Click()
    userName.Text = ""
    password.Text = ""
    password1.Text = ""
End Sub
Private Sub userName_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub
Private Sub password_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

Private Sub password1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        KeyAscii = 0
    End If
End Sub

