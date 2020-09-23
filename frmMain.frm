VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Shell About"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Display Dialog"
      Height          =   525
      Left            =   1740
      TabIndex        =   0
      Top             =   1350
      Width           =   1245
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Shell About Dialog, by Max Raskin, February 18 2000

'***

'This code shows how to show up the
'window's about dialog BUT with information
'on your application - icon, about etc..
'
'So, what is it good for me ?  I can
'do my own about dialog you ask
'
'The anwser is simple, as you might
'notice the window's about dialog shows up
'also system resources and windows version
'so it is really nice to know the current system
'resources that are left in a really quick way
'of showing this dialog:

'An API declaration to show show up the dialog
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

'The parameters of the API functions are:
'hwnd - the window handle of the window that will own the dialog (i.e. Me.hwnd)
'szApp - a string of the name of your application
'szOtherStuff - a string that showen in the middle of the dialog
               'can be anything you want - description or just some info
'hIcon - an icon that will be used for the dialog (i.e. Me.Icon or Picture1.Image)

'Here comes the code:
Private Sub Command1_Click()
    ShellAbout Me.hwnd, "ShellAbout", "This is the Shell About dialog box", Me.Icon
End Sub
