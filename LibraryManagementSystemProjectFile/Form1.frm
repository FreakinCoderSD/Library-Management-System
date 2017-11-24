VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPAC"
   ClientHeight    =   6315
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9630
   ClipControls    =   0   'False
   FillColor       =   &H0000FF00&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6315
   ScaleWidth      =   9630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton requiseBtn 
      BackColor       =   &H0000C000&
      Caption         =   "Send Requisition"
      Height          =   975
      Left            =   7800
      MaskColor       =   &H00E0E0E0&
      TabIndex        =   10
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton previousBtn 
      Caption         =   "<<"
      Height          =   375
      Left            =   8160
      TabIndex        =   9
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton nextBtn 
      Caption         =   ">>"
      Height          =   375
      Left            =   8160
      TabIndex        =   8
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Search Results"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Width           =   7335
      Begin VB.Label msg 
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
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   0
         Top             =   3360
         Width           =   6975
      End
      Begin VB.Label callNo 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   1920
         Width           =   4815
      End
      Begin VB.Label subject 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   1440
         Width           =   4815
      End
      Begin VB.Label author 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   4815
      End
      Begin VB.Label bookName 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   6975
      End
   End
   Begin VB.OptionButton byCallNo 
      Caption         =   "By Call Number"
      Height          =   375
      Left            =   3120
      TabIndex        =   6
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton byAuthor 
      Caption         =   "By Author"
      Height          =   375
      Left            =   1680
      TabIndex        =   5
      Top             =   1560
      Width           =   1215
   End
   Begin VB.OptionButton byName 
      Caption         =   "By Name"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton SearchBtn 
      Caption         =   "Search"
      Height          =   495
      Left            =   6480
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox searchBox 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label Label1 
      Caption         =   "Search A Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public con As New ADODB.Connection
Public rs1 As New ADODB.Recordset
Public rs2 As New ADODB.Recordset
Dim flag As Integer
Dim strCon As String
Dim sql As String
Dim sql2 As String

Public Function checkAvil()
Set rs2 = New ADODB.Recordset
If callNo.Caption <> "" Then
    sql2 = "select top 1 * from bookshelf where callNumber = '" & callNo.Caption & "' AND status = '1' "
    rs2.Open sql2, con, adOpenDynamic, adLockOptimistic, adCmdTxt
    If rs2.EOF And rs2.BOF Then
        requiseBtn.Enabled = False
        flag = 0
        msg.Caption = "No Copy of this Book In Stock"
    Else
        requiseBtn.Enabled = True
        msg.Caption = "Copy Available"
        flag = 1
    End If
End If
End Function
Public Function clear()
    searchBox.Text = ""
    bookName.Caption = ""
    callNo.Caption = ""
    subject.Caption = ""
    author.Caption = ""
    Set rs1 = Nothing
    Set rs2 = Nothing
    
End Function

Private Sub byAuthor_Click()
If byAuthor.Value = True Then
    searchBox.MaxLength = 0
End If

End Sub

Private Sub byCallNo_Click()
 If byCallNo.Value = True Then
    searchBox.MaxLength = 16
End If
If searchBox.MaxLength > 16 Then
    searchBox.Text = Left$(searchBox.Text, 16)
End If
End Sub

Private Sub byName_Click()
If byName.Value = True Then

searchBox.MaxLength = 0
End If
End Sub

Private Sub Form_Unload(cancel As Integer)
    Unload Me
End Sub

Private Sub nextBtn_Click()
rs1.MoveNext
If rs1.EOF = True Then
rs1.MoveFirst
    bookName.Caption = rs1(1)
    author.Caption = rs1(2)
    subject.Caption = rs1(3)
    callNo.Caption = rs1(5)
Else
    bookName.Caption = rs1(1)
    author.Caption = rs1(2)
    subject.Caption = rs1(3)
    callNo.Caption = rs1(5)
    Call checkAvil
End If
End Sub

Private Sub Form_Load()
    flag = 0
    requiseBtn.Enabled = False
    'create the database connection string
    strCon = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & App.Path & "\Database.mdb;Persist Security Info=False"
    nextBtn.Enabled = False
    previousBtn.Enabled = False
End Sub

Private Sub previousBtn_Click()
rs1.MovePrevious
If rs1.BOF Then
    rs1.MoveLast
    bookName.Caption = rs1(1)
    author.Caption = rs1(2)
    subject.Caption = rs1(3)
    callNo.Caption = rs1(5)
    Call checkAvil
Else
    bookName.Caption = rs1(1)
    author.Caption = rs1(2)
    subject.Caption = rs1(3)
    callNo.Caption = rs1(5)
    Call checkAvil
End If
End Sub

Private Sub requiseBtn_Click()
If flag = 1 Then
     Form1.Enabled = False
        
    sendRequisition.Show
End If

End Sub

Private Sub searchBox_Change()
flag = 0
If byCallNo.Value = True Then
    searchBox.MaxLength = 16
Else
    searchBox.MaxLength = 0
End If

End Sub

Private Sub searchBox_LostFocus()
 searchBox.Text = RemoveExtraSpaces(searchBox.Text)
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
Private Sub SearchBtn_Click()
If byName.Value = True Then
    If searchBox.Text = "" Then
        MsgBox "Enter The Name of The Book"
        searchBox.SetFocus
    Else
        Set con = New ADODB.Connection
        con.Open strCon 'Open connection
        Set rs1 = New ADODB.Recordset
        sql = "SELECT * FROM bookDetails WHERE bookName LIKE '%" & searchBox.Text & "%'"
        rs1.Open sql, con, adOpenDynamic, adLockOptimistic, adCmdText
       If rs1.EOF And rs1.BOF Then
            bookName.Caption = "No record found"
            author.Caption = ""
            subject.Caption = ""
            callNo.Caption = ""
            msg.Caption = ""
            previousBtn.Enabled = False
            nextBtn.Enabled = False
        Else
        
            rs1.MoveFirst
            bookName.Caption = rs1(1)
            author.Caption = rs1(2)
            subject.Caption = rs1(3)
            callNo.Caption = rs1(5)
            previousBtn.Enabled = True
            nextBtn.Enabled = True
        End If
                
    End If
    
ElseIf byCallNo.Value = True Then
    searchBox.Text = Left$(searchBox.Text, 16)
    
    If searchBox.Text = "" Then
        MsgBox "Enter the call number of the book"
        searchBox.SetFocus
    Else
        Set con = New ADODB.Connection
        con.Open strCon 'Open connection
        Set rs1 = New ADODB.Recordset
        sql = "SELECT * FROM bookDetails WHERE callNumber = '" & searchBox.Text & "'"
        rs1.Open sql, con, adOpenDynamic, adLockOptimistic, adCmdText
        If rs1.EOF And rs1.BOF Then
            bookName.Caption = "No record found"
            author.Caption = ""
            subject.Caption = ""
            callNo.Caption = ""
            msg.Caption = ""
             previousBtn.Enabled = False
            nextBtn.Enabled = False
        Else
        
            rs1.MoveFirst
            bookName.Caption = rs1(1)
            author.Caption = rs1(2)
            subject.Caption = rs1(3)
            callNo.Caption = rs1(5)
            previousBtn.Enabled = True
            nextBtn.Enabled = True
        End If
    End If
        
ElseIf byAuthor.Value = True Then
    If searchBox.Text = "" Then
        MsgBox "Enter the author name of the book"
        searchBox.SetFocus
    Else
        Set con = New ADODB.Connection
        con.Open strCon 'Open connection
        Set rs1 = New ADODB.Recordset
        sql = "SELECT * FROM bookDetails WHERE authorName LIKE '%" & searchBox.Text & "%'"
        rs1.Open sql, con, adOpenDynamic, adLockOptimistic, adCmdText
        If rs1.EOF And rs1.BOF Then
            bookName.Caption = "No record found"
            author.Caption = ""
            subject.Caption = ""
            callNo.Caption = ""
            msg.Caption = ""
             previousBtn.Enabled = False
            nextBtn.Enabled = False
        Else
        
            rs1.MoveFirst
            bookName.Caption = rs1(1)
            author.Caption = rs1(2)
            subject.Caption = rs1(3)
            callNo.Caption = rs1(5)
             previousBtn.Enabled = True
            nextBtn.Enabled = True
        End If
    End If
End If
Call checkAvil

End Sub


