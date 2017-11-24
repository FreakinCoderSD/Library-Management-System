VERSION 5.00
Begin VB.Form addBook 
   Caption         =   "Add A New Book"
   ClientHeight    =   5100
   ClientLeft      =   120
   ClientTop       =   525
   ClientWidth     =   7065
   LinkTopic       =   "Form2"
   ScaleHeight     =   5100
   ScaleWidth      =   7065
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3600
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1800
      TabIndex        =   9
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   15
         Top             =   2280
         Width           =   4575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "add"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   4080
         Width           =   1575
      End
      Begin VB.CommandButton generate 
         Caption         =   "Generate"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5280
         TabIndex        =   11
         Top             =   3480
         Width           =   1215
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   10
         Top             =   1680
         Width           =   4575
      End
      Begin VB.TextBox Text5 
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
         Height          =   495
         Left            =   1680
         TabIndex        =   8
         Top             =   3480
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   7
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1680
         TabIndex        =   6
         Top             =   480
         Width           =   4575
      End
      Begin VB.Label Label6 
         Caption         =   "Subject"
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
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.Label Label5 
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   3600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "No Of Copies"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Entry Date"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Author Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
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
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
End
Attribute VB_Name = "addBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim con As New ADODB.Connection
Dim str As String
Private Sub Command1_Click()
Set con = New ADODB.Connection
con.Open str
Set rs1 = New ADODB.Recordset
Set rs2 = New ADODB.Recordset

Dim copy As Integer
Dim i As Integer
If Text5.Text = "" Then
    MsgBox ("Please fill all the field first")
    If Text1.Text = "" Then
    Text1.SetFocus
    
    ElseIf Text2.Text = "" Then
    Text2.SetFocus
    
    ElseIf Text3.Text = "" Then
    Text3.SetFocus
    
    ElseIf Text4.Text = "" Then
    Text4.SetFocus
      
End If
Else
 
 rs1.Open "select * from bookDetails WHERE callNumber = '" & Text5.Text & "'", con, adOpenDynamic, adLockOptimistic, adCmdText
 rs2.Open "select * from bookshelf", con, adOpenDynamic, adLockOptimistic, adCmdText
    copy = Val(Text4.Text)
    If rs1.EOF And rs1.BOF Then
    
    rs1.AddNew
    rs1(1) = Text1.Text
    rs1(2) = Text2.Text
    rs1(3) = Text6.Text
    rs1(4) = Text3.Text
    rs1(5) = Text5.Text
    rs1.update
    
    i = 1
    For i = 1 To copy
    rs2.AddNew
    rs2(1) = Text1.Text
    rs2(2) = Text5.Text
    rs2(3) = Text5.Text + "_c" + CStr(Format(i, "000"))
    rs2(4) = "1"
    rs2.update
    
    Next
    MsgBox "The new book is added successfully", vbOKOnly, "Book added successfully"
    Call Command2_Click
    
    Else
        MsgBox "This book is already in the database", vbOKOnly, "Book already added"
    End If
    
End If

End Sub


Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text1.SetFocus

End Sub

Private Sub Form_Load()
Admin.Enabled = False
str = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=TRUE"
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text6.Text = ""
Text5.Text = ""
End Sub

Private Sub Form_Unload(cancel As Integer)
    Admin.Enabled = True
    Admin.Show
End Sub

Private Sub generate_Click()
Dim callNo As String

If Text1.Text = "" Then
MsgBox ("Enter the Book Name.")
Text1.SetFocus

ElseIf Text2.Text = "" Then
MsgBox ("Enter the Author Name")
Text2.SetFocus

ElseIf Text6.Text = "" Then
MsgBox ("Enter the Subject")
Text6.SetFocus

ElseIf Text3.Text = "" Then
MsgBox ("Enter the Date")
Text3.SetFocus

ElseIf Text4.Text = "" Then
MsgBox ("Enter the No Of Copies")
Text4.SetFocus

 ElseIf Text4.Text = "0" Then
        MsgBox "You must insert atleast one copy.", vbOKOnly, "Insert atleast one copy"
        Text4.SetFocus
        
ElseIf Len(Text1.Text) < 5 Then
    MsgBox "Book name contain atleast 5 character ", vbCritical, "Invalid book name"
    Text1.SetFocus
    
ElseIf Len(Text2.Text) < 4 Then
    MsgBox "Author name contain atleast 4 character ", vbCritical, "Invalid book name"
    Text1.SetFocus
Else
callNo = Left$(Text2.Text, 3) + "." + Right$(Text1.Text, 2) + ":" + CStr(Format(Len(Text1.Text), "00")) + "\" + Mid$(Text1.Text, 2, 3) + "." + CStr(Format(Len(Text2.Text), "00"))
Text5.Text = callNo
End If

End Sub

Private Sub Text1_LostFocus()
    Text1.Text = RemoveExtraSpaces(Text1.Text)
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = RemoveExtraSpaces(Text2.Text)
End Sub
    
Public Function RemoveExtraSpaces(str As String) As String
    
    str = Trim$(str)
    
    Dim L As Integer, i As Integer
    Dim S As String
    Dim Prev_char As String
    
    S = ""
    
    L = Len(str)
    i = 1
    Do
        Prev_char = Mid$(str, i, 1)
        i = i + 1
        
        S = S + Prev_char
        If Prev_char = " " Then
            Do While (i < L) And (Mid$(str, i, 1) = " ")
                i = i + 1
            Loop
        End If
        
    Loop Until i > L
    
    str = S
    RemoveExtraSpaces = S
End Function

Private Sub Text3_GotFocus()
Text3.Text = DateValue(Now)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
