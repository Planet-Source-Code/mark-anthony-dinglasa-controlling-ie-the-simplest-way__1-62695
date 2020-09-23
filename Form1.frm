VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tamers IE Control 1.0"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   5400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   615
      Left            =   3840
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   615
      Left            =   2640
      TabIndex        =   6
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdForward 
      Caption         =   "Forward"
      Height          =   615
      Left            =   1560
      TabIndex        =   5
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "Go"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
   End
   Begin VB.TextBox txtSite 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\"
      Top             =   1560
      Width           =   3615
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   855
      Left            =   0
      Top             =   2160
      Width           =   5415
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3840
      Picture         =   "Form1.frx":1272
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Controlling Internet Explorer !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   360
      TabIndex        =   3
      Top             =   0
      Width           =   4680
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      FillColor       =   &H00FFC0C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Input website to view Or You can browse folders in any Drives you specify !"
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
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   3495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   1575
      Left            =   0
      Top             =   480
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''
' First just add a Microsoft Internet Control
' component which we will use to show up the
' Windows Internet Explorer in his own window
' then on there we will control its functionality
''''''''''''''' Have Fun ! '''''''''''''''''''''''

'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
''''''''''''''' Beginners Only '''''''''''''''''''
' Here is how to add a component
'1) Right Click on the Left part of the Visual
'   Basic Design Time where you usually found the other controls
'   such as the label, textbox, command button and etc.
'2) Select components
'3) Finally scroll down until you can find the
'   Microsoft Internet Control in the Listbox
'   under the Controls Tab
'4) Click apply/Ok....Respectively ! Have Fun !
''''''''''''''''''''''''''''''''''''''''''''''''''
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

'This uses withevents to make "Iview" variable act
'as the WebBrowser
Private WithEvents Iview As SHDocVwCtl.WebBrowser_V1
Attribute Iview.VB_VarHelpID = -1

Private Sub cmdBack_Click()
On Error Resume Next 'use for trapping an expected errors
Iview.GoBack
txtSite.Text = Iview.LocationURL
End Sub

Private Sub cmdForward_Click()
On Error Resume Next
Iview.GoForward
txtSite.Text = Iview.LocationURL
End Sub

Private Sub cmdGo_Click()
Iview.Visible = True 'make internet explorer visible
Iview.Navigate txtSite.Text 'after it is visible it will then navigate to the specified site of drives/folders
End Sub

Private Sub cmdRefresh_Click()
    Iview.Refresh
End Sub

Private Sub cmdSearch_Click()
    Iview.GoSearch
End Sub

Private Sub Form_Load() 'Setting Iview to Internet Explorer
Set Iview = GetObject("", "InternetExplorer.Application")
End Sub

Private Sub Form_Unload(Cancel As Integer)
Iview.Visible = False 'Hiding Internet Explorer
Set Iview = Nothing   'Clearing "Iview" from Memory
End Sub

Private Sub txtSite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then '13 is equivalent to "Enter Key"
        Call cmdGo_Click
    End If
End Sub

'''''''''''''''''' Advance Users Read Here''''''''''''''
'For those who are Advance users just add "AlwaysOnTop"
'Code using Windows API to make it always on the top on
'all applications running.I will not include it in here
'for this program is beginners purpose only !
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'xxxxxxxxxxxxxxxxxxx 09/27/2005 Tuesday xxxxxxxxxxxxxxxxx
' For comments and suggestions just email
' mark_anthony_dinglasa@ yahoo.com
' also visit www.geocities.com/mark_anthony_dinglasa/2003
'xxxxxxxxxxxxxxxxxxx Have a nice day !xxxxxxxxxxxxxxxxxxx
