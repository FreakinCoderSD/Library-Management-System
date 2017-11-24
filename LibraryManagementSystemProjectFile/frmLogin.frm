VERSION 5.00
Begin VB.Form frmLogin 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1575
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   930.562
   ScaleMode       =   0  'User
   ScaleWidth      =   3647.805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1320
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   495
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public con As New ADODB.Connection
Public rs As New ADODB.Recordset
Dim sql As String
Dim constr As String
Dim passwordcheck As Boolean
Public loggedin As Boolean
Private Sub cmdCancel_Click()
    Unload Me
    loggedin = False
End Sub

Private Sub cmdOK_Click()
    Set con = New ADODB.Connection
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    con.Open constr
    Set rs = New ADODB.Recordset
    sql = "SELECT * FROM admin WHERE userName = '" & txtUserName.Text & "' AND password = '" & txtPassword.Text & "'"
    rs.Open sql, con, adOpenDynamic, adLockOptimistic
    If rs.EOF And rs.BOF Then
        MsgBox "User name OR password is incorrect", vbCritical, "Incorrect login details"
        passwordcheck = False
        txtUserName = ""
        txtPassword = ""
        
    Else
        Admin.Enabled = True
        Admin.Show
        passwordcheck = True
        Unload Me
        loggedin = True
    End If
    
    
End Sub
    
Private Sub Form_Load()
  passwordcheck = False
    loggedin = False
End Sub
Private Sub Form_Unload(cancel As Integer)
    If rs.State = 1 Then
    rs.Close
    con.Close
    End If
    If passwordcheck = False Then
        Unload Admin
    End If
End Sub

