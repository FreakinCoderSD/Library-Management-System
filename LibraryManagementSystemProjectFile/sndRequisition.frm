VERSION 5.00
Begin VB.Form sendRequisition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Send Requisition"
   ClientHeight    =   3960
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox rollBox 
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
      Height          =   405
      Left            =   1680
      TabIndex        =   14
      Top             =   840
      Width           =   4215
   End
   Begin VB.TextBox nameBox 
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
      Height          =   405
      Left            =   1680
      TabIndex        =   13
      Top             =   120
      Width           =   4215
   End
   Begin VB.TextBox subject 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   10
      Top             =   3240
      Width           =   5655
   End
   Begin VB.TextBox callNo 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   9
      Top             =   2760
      Width           =   5655
   End
   Begin VB.TextBox author 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1440
      TabIndex        =   8
      Top             =   2280
      Width           =   5655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   7215
      Begin VB.TextBox bookName 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1320
         TabIndex        =   7
         Top             =   360
         Width           =   5655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Subject"
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
         Top             =   1920
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   5
         Top             =   1440
         Width           =   1110
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   4
         Top             =   960
         Width           =   1170
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   3
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   6240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Roll Number"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   12
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   675
   End
End
Attribute VB_Name = "sendRequisition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs3 As New ADODB.Recordset
Dim sql As String
Dim today As String
Dim time As String
Dim intResponce As Integer

Private Sub CancelButton_Click()
Unload Me
Form1.Enabled = True
End Sub

Private Sub Form_Load()
bookName.Text = Form1.rs1(1)
author.Text = Form1.rs1(2)
callNo.Text = Form1.rs1(5)
subject.Text = Form1.rs1(3)
End Sub

Private Sub Form_Unload(cancel As Integer)
Form1.Enabled = True
End Sub
Private Sub OKButton_Click()
If nameBox = "" Then
    MsgBox "Enter Your Name"
    nameBox.SetFocus
ElseIf rollBox = "" Then
    MsgBox "Enter Your Roll"
    rollBox.SetFocus
Else
    intResponce = MsgBox("Are you sure to send requisition for this book", vbYesNoCancel, "Confirm requisition")
    If intResponce = vbYes Then
        today = CStr(DateValue(Now))
        time = CStr(TimeValue(Now))
        Set rs3 = New ADODB.Recordset
        sql = "SELECT * FROM requisition"
        rs3.Open sql, Form1.con, adOpenDynamic, adLockOptimistic
        rs3.AddNew
        rs3.Fields(1) = bookName.Text
        rs3.Fields(2) = Form1.rs2(3)
            Form1.rs2.Fields(4) = "0"
            Form1.rs2.update
        rs3.Fields(3) = DateValue(Now)
        rs3.Fields(4) = TimeValue(Now)
        rs3.Fields(5) = nameBox.Text
        rs3.Fields(6) = rollBox.Text
        rs3.update
        'Set rs3 = Nothing
        rs3.Close
        Form1.checkAvil
        Unload Me
        Form1.Enabled = True

        Form1.Show
        MsgBox "Your book requisition has been submitted successfully. Librarian will response soon.", vbOKOnly, "Requisition successful"
        
    ElseIf intResponce = vbNo Then
        Unload Me
        
    ElseIf vbCancel Then
        
    End If
        
End If

End Sub
