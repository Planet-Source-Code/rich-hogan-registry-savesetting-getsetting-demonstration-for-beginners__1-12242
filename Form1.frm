VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "GetSetting/SaveSetting Demo"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4305
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4305
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "Check :    On/Off"
      Height          =   300
      Left            =   165
      TabIndex        =   3
      Top             =   2175
      Value           =   1  'Checked
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Arial (Now and on next load)"
      Height          =   285
      Left            =   105
      TabIndex        =   2
      Top             =   1740
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000004&
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Text            =   "Choose Font"
      Top             =   150
      Width           =   3960
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":27A2
      Top             =   480
      Width           =   3960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "www.rockandice.co.uk"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   255
      MouseIcon       =   "Form1.frx":2877
      MousePointer    =   10  'Up Arrow
      TabIndex        =   4
      Top             =   2790
      Width           =   3615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////
'      Rich Hogan.       /
'   rock.ice@bigfoot.com /
'                        /
'   www.rockandice.co.uk /
'                        /
'                        /
'                        /
'/////////////////////////



Private Sub Check1_Click()
'save users selection on click
SaveSetting App.Title, "Settings", "Checkbox", Check1.Value
End Sub

Private Sub Combo1_Click()
SaveSetting App.Title, "Settings", "Font", Combo1.Text
   Text1.Font = Combo1.Text
End Sub

Private Sub Command1_Click()
'Just resets font to Arial
SaveSetting App.Title, "Settings", "Font", "Arial"
Text1.FontName = GetSetting(App.Title, "Settings", "Font")
End Sub

Private Sub Form_Load()
On Error GoTo Regerror 'See err: Notes
'Load combo's list
Combo1.AddItem "Symbol"
Combo1.AddItem "Tahoma"
Combo1.AddItem "Westminster"
Combo1.AddItem "MS Sans Serif"
Combo1.AddItem "Arial"
Combo1.AddItem "System"
Combo1.AddItem "Wingdings"
'///////////////////////////////////////

'Get setting from registry for check1:
Check1.Value = GetSetting(App.Title, "Settings", "Checkbox")

'Get setting from registry for Text1 Font style
Text1.FontName = GetSetting(App.Title, "Settings", "Font")

'///////////////////////////////////////
'Play with the font values a little:
If GetSetting(App.Title, "Settings", "Font") = "Westminster" Then
   MsgBox "Westminster font was the chosen font"
Else
   MsgBox "Text1 font is not Westminster!"
End If
'///////////////////////////////////////
Regerror:
'On error load this code to clear an invalid entry.
'remember to close app then rem out this line again or will not work! :
'SaveSetting App.Title, "Settings", "Font", "Arial"
'Text1.FontName = GetSetting(App.Title, "Settings", "Font")
End Sub

Private Sub Label1_Click()
Shell ("Start http://www.rockandice.co.uk")
End Sub
