VERSION 5.00
Begin VB.Form fDemo 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Tooltip Demo"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmrHasNoToolTip 
      Left            =   2025
      Top             =   2325
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   450
      Index           =   1
      Left            =   195
      TabIndex        =   8
      ToolTipText     =   "This is another Command Button 3"
      Top             =   2325
      Width           =   1215
   End
   Begin VB.CommandButton btExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2985
      TabIndex        =   7
      Top             =   2325
      Width           =   750
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   780
      Left            =   2130
      TabIndex        =   4
      ToolTipText     =   "This is Frame Number 1|holding one Option Box  "
      Top             =   1275
      Width           =   1620
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   195
         Left            =   315
         TabIndex        =   5
         ToolTipText     =   "This is Option 1"
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2130
      MaxLength       =   16
      TabIndex        =   3
      Text            =   "Text1"
      ToolTipText     =   "This is Textbox 1 - Enter some text|and hover mouse again"
      Top             =   225
      Width           =   1620
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   450
      Index           =   0
      Left            =   195
      TabIndex        =   2
      ToolTipText     =   "This is Command Button 3"
      Top             =   1605
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   450
      Left            =   195
      TabIndex        =   1
      ToolTipText     =   "This is Command Button 2"
      Top             =   900
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   450
      Left            =   195
      TabIndex        =   0
      ToolTipText     =   "This is Command Button 1"
      Top             =   180
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Label1"
      Height          =   255
      Left            =   2130
      TabIndex        =   6
      ToolTipText     =   "This is Label1 (sorry, no hWnd - no Custom ToolTip)"
      Top             =   780
      Width           =   1620
   End
End
Attribute VB_Name = "fDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Note that this collection is automatically destroyed when this form is
'unloaded causing an avalanche effect such that all class instances are
'destroyed when this collection is destroyed (it having the only reference
'to each class instance), and on Class_Terminate all created tool windows
'are destroyed, so I'm pretty sure there is no memory leak :-)
Private Tooltips        As New Collection 'keeping references to all tooltip class instances

Private Sub btExit_Click()

    Unload Me

End Sub

Private Sub Form_Load()

  'This is called once on Form Load and creates all relevant tooltip windows
  'The original ToolTipText property is used to fill these windows with life.
  'The vertical bar | is used as line break character, however line breaks
  'are not possible unless the tooltip also has a title, you may or may not
  'include a tooltip title (see params for .Create)
  'Individual back- and forecolors are also possible as well as an assortment
  'of Icons to be displayed in the tooltip and individual hover- and popup-times.
  'The .Create function returns the hWnd of the created tooltip window
  'or zero if unsuccessful.

  Dim Tooltip   As cToolTip
  Dim Control   As Control
  Dim CollKey   As String
  Dim e         As Long

    For Each Control In Controls 'cycle thru all controls
        With Control
            On Error Resume Next 'in case the control has no tooltiptext property
                CollKey = .ToolTipText 'try to access that property
                e = Err 'save error
            On Error GoTo 0
            If e = 0 Then 'the control has a tooltiptext property
                If Len(Trim$(.ToolTipText)) Then 'use that to create the custom tooltip
                    CollKey = .Name
                    On Error Resume Next 'in case control is not in an array of controls and therefore has no index property
                        CollKey = CollKey & "(" & .Index & ")"
                    On Error GoTo 0
                    Set Tooltip = New cToolTip
                    If Tooltip.Create(Control, Trim$(.ToolTipText), TTBalloonAlways, (TypeName(Control) = "TextBox"), TTIconInfo, CollKey) Then
                        Tooltips.Add Tooltip, CollKey 'to keep a reference to the current tool tip class instance (prevent it from being destroyed)
                        .ToolTipText = vbNullString 'kill tooltiptext so we don't get two tips
                    End If
                End If
            End If
        End With 'CONTROL
    Next Control

    'and one indvidual tooltip
    Set Tooltip = New cToolTip
    Tooltip.Create btExit, "Click on this button|to close application", TTBalloonIfActive, False, TTIconError, "Exit", vbBlue, vbCyan, 1500, 20000
    Tooltips.Add Tooltip, btExit.Name 'don't forget to keep it

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  'this demonstrates how a tooltip could be altered at runtime
  'uses late binding but who cares? (only happens once during mouse move after the text has changed)

    If Text1.DataChanged Then
        Text1.DataChanged = False
        With Tooltips(Text1.Name)  'finds reference to Text1 tooltip class instance
            .Create Text1, .InitialText & "||Text has changed:|" & """" & Text1 & """", .Style, .Centered, TTIconWarning, .InitialTitle
        End With 'TOOLTIPS("TEXT1")'TOOLTIPS(TEXT1.NAME)
    End If

End Sub

':) Ulli's VB Code Formatter V2.15.4 (2003-Jan-09 10:59) 8 + 67 = 75 Lines
