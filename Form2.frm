VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4656
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8292
   LinkTopic       =   "Form2"
   ScaleHeight     =   4656
   ScaleWidth      =   8292
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Create Field Object"
      Height          =   372
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   2052
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Fields Values"
      Height          =   372
      Left            =   120
      TabIndex        =   3
      Top             =   1920
      Width           =   2052
   End
   Begin VB.TextBox Text1 
      Height          =   288
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   3492
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   372
      Left            =   7080
      TabIndex        =   0
      Top             =   4080
      Width           =   972
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   252
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   732
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim objRec As ADODB.Recordset
Dim objFld As ADODB.Field
Dim objFlds As ADODB.Fields
Dim obj As ADODB.Command

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
'Show values of all fields for current record
Dim i As Integer
Dim str As String
On Error Resume Next
'Dim objRecs As ADODB.Recordset
'Set objRecs = New ADODB.Recordset
'objRecs.Open "Publishers", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Program Files\Microsoft Visual Studio\VB98\Biblio.mdb;Persist Security Info=False"
str = ""

For i = 0 To objRec.Fields.Count - 1
    str = str & objRec.Fields.Item(i).Value & vbCrLf
Next
MsgBox str
End Sub

Private Sub Command3_Click()
Set objFlds = objRec.Fields
For Each objFld In objFlds
    MsgBox objFld.Name
Next
Dim i As Integer
For i = 0 To objFlds.Count - 1
    MsgBox objFlds(i).Name
Next
Set objFlds = Nothing
End Sub

Private Sub Form_Load()
Set objRec = Form1.Adodc1.Recordset
Text1.Text = objRec("Name").Value

End Sub

Private Sub Form_Unload(Cancel As Integer)
objRec("Name").Value = Text1.Text
End Sub
