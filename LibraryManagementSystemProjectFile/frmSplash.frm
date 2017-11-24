VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3975
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7125
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7185
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   360
         Picture         =   "frmSplash.frx":000C
         ScaleHeight     =   810
         ScaleWidth      =   810
         TabIndex        =   4
         Top             =   720
         Width           =   810
      End
      Begin VB.Timer Timer1 
         Interval        =   2000
         Left            =   6120
         Top             =   960
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6360
         TabIndex        =   5
         Top             =   2640
         Width           =   450
      End
      Begin VB.Label lblLicenseTo 
         Caption         =   "Designed And Developed By :  Payel De, Sayan Dasgupta, Subhashis Pal"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   3480
         Width           =   6855
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "v"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   1
         Top             =   2640
         Width           =   195
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Pro Librarian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   2250
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim flag As Integer '' two check the admin already exists or not
Dim con As ADODB.Connection
Dim rs As ADODB.Recordset
Dim sql As String
Dim constr As String

Private Sub Form_Load()
    Set con = New ADODB.Connection
    constr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    con.Open constr
    Set rs = New ADODB.Recordset
    sql = "SELECT * FROM admin "
    rs.Open sql, con, adOpenDynamic, adLockOptimistic
    If rs.EOF And rs.BOF Then
        flag = 1
        
    Else
        flag = 0
    End If
    
End Sub

Private Sub lblVersion_Click()

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
Admin.Show
If flag = 1 Then
    frmSignUp.Show
Else
    frmLogin.Show
End If
Unload Me
End Sub
