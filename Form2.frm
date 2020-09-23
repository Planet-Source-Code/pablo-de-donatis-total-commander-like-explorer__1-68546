VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   570
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   570
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   720
      TabIndex        =   0
      Text            =   "*.*"
      Top             =   120
      Width           =   1215
   End
   Begin Project1.dcButton dcButton1 
      Default         =   -1  'True
      Height          =   300
      Left            =   2040
      TabIndex        =   2
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      BackColor       =   14995922
      ButtonShape     =   3
      ButtonStyle     =   9
      Caption         =   ""
      Effects         =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicAlign        =   0
      PicNormal       =   "Form2.frx":0000
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin Project1.dcButton dcButton2 
      Cancel          =   -1  'True
      Height          =   300
      Left            =   2760
      TabIndex        =   3
      Top             =   120
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   529
      BackColor       =   14995922
      ButtonShape     =   3
      ButtonStyle     =   9
      Caption         =   ""
      Effects         =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicNormal       =   "Form2.frx":015A
      PicSizeH        =   16
      PicSizeW        =   16
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Filter:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Left            =   120
      TabIndex        =   1
      Top             =   150
      Width           =   450
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub dcButton1_Click()

If Text1.Text <> Form1.StatusBar1.Panels(2).Text Then
Form1.StatusBar1.Panels(2).Text = Text1.Text
Form1.StatusBar1.Tag = "Ok"
Unload Me
Else
Call dcButton2_Click
End If
End Sub

Private Sub dcButton2_Click()
Form1.StatusBar1.Tag = "Cancel"
Unload Me
End Sub

Private Sub Form_Load()
Text1.SelStart = Len(Text1.Text)
End Sub
