VERSION 5.00
Begin VB.Form frmEncryption 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Encryption Form"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text5 
      Height          =   3255
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Text            =   "frmEncryption.frx":0000
      Top             =   2400
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   3600
      TabIndex        =   6
      Text            =   "Text4"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   5760
      Width           =   5535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Decrypt"
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "frmEncryption.frx":0006
      Top             =   1320
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Encrypt It"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   480
      Width           =   5535
   End
   Begin VB.Label Label1 
      Caption         =   "How Many Times to Encrypt"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Message To Encrypt"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmEncryption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim HMT As Long
    HMT = Text4.Text
    Text2.Text = Encrypt(Text1.Text)
    Text5.Text = ""
    For i = 1 To HMT
        Text5.Text = Text5.Text & i & ": "
        Text5.Text = Text5.Text & Encrypt(Text1.Text)
        Text5.Text = Text5.Text & vbCrLf
        DoEvents: DoEvents: DoEvents: DoEvents
    Next i
End Sub

Private Sub Command3_Click()
    Text3.Text = Decrypt(Text2.Text)
End Sub

Private Sub Form_Load()
Text1.Text = "Encrypt Me"
Text2.Text = ""
Text3.Text = ""
Text4.Text = "15"
Text5.Text = ""
End Sub

