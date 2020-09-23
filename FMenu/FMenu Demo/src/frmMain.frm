VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FMenu Demo"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3480
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   3480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLD 
      Caption         =   "Load and display"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmMAin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private cFM As clsFMenu
' There's no point in using events for FMenu. The
'   internal scripting system is meant to do all of the
'   work. Basically, either the menu loads or it doesn't.

Private Sub cmdLD_Click()
    Set cFM = New clsFMenu
        With cFM
            .OwnerHWND = Me.hwnd     'Required for TrackPopupMenu()
            Call .LoadMenus(App.Path & "\" & "Demo Menu.ini")
            Call .ShowMenu((Me.Left / Screen.TwipsPerPixelX), (Me.Top / Screen.TwipsPerPixelY))
        End With
    Set cFM = Nothing
End Sub

'Now how difficult was that for you, sir jbooks? :p
