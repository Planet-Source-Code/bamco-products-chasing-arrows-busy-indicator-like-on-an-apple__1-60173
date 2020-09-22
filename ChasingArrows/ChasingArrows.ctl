VERSION 5.00
Begin VB.UserControl ChasingArrows 
   BackColor       =   &H80000001&
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   EditAtDesignTime=   -1  'True
   PaletteMode     =   0  'Halftone
   ScaleHeight     =   16
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   16
   ToolboxBitmap   =   "ChasingArrows.ctx":0000
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1920
      Top             =   2400
   End
   Begin VB.Image Vic 
      Height          =   240
      Left            =   0
      Picture         =   "ChasingArrows.ctx":0312
      Tag             =   "1"
      Top             =   0
      Width           =   240
   End
   Begin VB.Image Image8 
      Height          =   240
      Left            =   2760
      Picture         =   "ChasingArrows.ctx":0656
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   2520
      Picture         =   "ChasingArrows.ctx":099A
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   2280
      Picture         =   "ChasingArrows.ctx":0CDE
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   2040
      Picture         =   "ChasingArrows.ctx":1022
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   1800
      Picture         =   "ChasingArrows.ctx":1366
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1560
      Picture         =   "ChasingArrows.ctx":16AA
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1320
      Picture         =   "ChasingArrows.ctx":19EE
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1080
      Picture         =   "ChasingArrows.ctx":1D32
      Top             =   1200
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ChasingArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Sub Timer1_Timer()
Select Case Vic.Tag

    Case 0
        Let Vic.Tag = 1
        Let Vic.Picture = Image1.Picture
    Case 1
        Let Vic.Tag = 2
        Let Vic.Picture = Image2.Picture
    Case 2
        Let Vic.Tag = 3
        Let Vic.Picture = Image3.Picture
    Case 3
        Let Vic.Tag = 4
        Let Vic.Picture = Image4.Picture
    Case 4
        Let Vic.Tag = 5
        Let Vic.Picture = Image5.Picture
    Case 5
        Let Vic.Tag = 6
        Let Vic.Picture = Image6.Picture
    Case 6
        Let Vic.Tag = 7
        Let Vic.Picture = Image7.Picture
    Case 7
        Let Vic.Tag = 8
        Let Vic.Picture = Image8.Picture
    Case 8
        Let Vic.Tag = 1
        Let Vic.Picture = Image1.Picture

End Select
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'If the control's container is in design mode, turn off the timer, which
'will cause the control to stop working.
Timer1.Enabled = Ambient.UserMode

'Get propertys from property bag.
Enabled = PropBag.ReadProperty("Enabled", True)
Timer1.Enabled = Ambient.UserMode

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'Save propertys in property bag.
PropBag.WriteProperty "Enabled", Enabled, True

End Sub

Public Property Get Enabled() As Boolean
'Enabled property is stored in Enabled property control.
Enabled = UserControl.Enabled

End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
'New Enabled value is passed to control object.
UserControl.Enabled = NewValue
UserControl.PropertyChanged "Enabled"

'Depending on control's condition - active or not - should modify object lblText and
'arrows on control.
Select Case NewValue
Case Is = True
 Timer1.Enabled = True
Case Is = False
 Timer1.Enabled = False
End Select

End Property
