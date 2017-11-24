VERSION 5.00
Begin VB.Form chngPassword 
   Caption         =   "Change Pasword"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   5415
   LinkTopic       =   "Form2"
   ScaleHeight     =   3645
   ScaleWidth      =   5415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton clear 
      Caption         =   "Clear"
      Height          =   375
      Left            =   1560
      TabIndex        =   10
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton saveBtn 
      Caption         =   "Save"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox newPassword 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   7
      Top             =   1680
      Width           =   2655
   End
   Begin VB.TextBox newUserName 
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
      Height          =   330
      Left            =   2280
      MaxLength       =   16
      TabIndex        =   6
      Top             =   1200
      Width           =   2655
   End
   Begin VB.TextBox oldPassword 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   16
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox oldUSerName 
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2280
      MaxLength       =   16
      TabIndex        =   0
      Top             =   195
      Width           =   2655
   End
   Begin VB.Label msg 
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
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   4815
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter Old Password"
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
      TabIndex        =   4
      Top             =   720
      Width           =   1770
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Enter New User Name"
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
      TabIndex        =   3
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter New Password"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   1860
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter User Name"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1545
   End
End
Attribute VB_Name = "chngPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sql As String
Dim con As ADODB.Connection
Dim rs As New ADODB.Recordset
Dim constr As String

Private Sub cancel_Click()
Unload Me
Admin.Enabled = True
Admin.Show
End Sub

Private Sub Form_Load()
    Admin.Enabled = False
    
End Sub

Private Sub Form_Unload(cancel As Integer)
Admin.Enabled = True
Admin.Show
End Sub

Private Sub saveBtn_Click()
    Set con = New ADODB.Connection
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    con.Open constr
    If Not oldUSerName.Text = "" And Not oldPassword.Text = "" And Not newUserName.Text = "" And Not newPassword.Text = "" Then
        If Len(oldUSerName.Text) < 5 Then
            msg.Caption = "User name must contain atleast 5 character. "
            msg.ForeColor = vbRed
        ElseIf Len(newUserName.Text) < 5 Then
            msg.Caption = "User name must contain atleast 5 character. "
            msg.ForeColor = vbRed
        ElseIf Len(newPassword.Text) < 5 Then
            msg.Caption = "Password must contain atleast 5 character "
            msg.ForeColor = vbRed
        ElseIf Len(oldPassword.Text) < 5 Then
            msg.Caption = "Password must contain atleast 5 character "
            msg.ForeColor = vbRed
        Else
            Set rs = New ADODB.Recordset
            sql = "SELECT * FROM admin WHERE userName = '" & oldUSerName & "' AND password = '" & oldPassword & "'"
            rs.Open sql, con, adOpenDynamic, adLockOptimistic
            If Not rs.EOF And Not rs.BOF Then
                rs.Fields(1) = newUserName.Text
                rs.Fields(2) = newPassword.Text
                rs.Fields(3) = DateValue(Now)
                rs.update
                MsgBox "User name and password updated successfully ", vbOKOnly, "Successfully updated"
                Unload Me
                Admin.Enabled = True
                Admin.Show
                
            
            Else
                msg.Caption = "Please enter the old user name and password correctly."
                oldUSerName.SetFocus
            End If
        End If
    End If
End Sub
