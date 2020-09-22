VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Strip tags and Submit to Webform"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   405
   ClientWidth     =   10245
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      Height          =   375
      Left            =   9120
      TabIndex        =   13
      Top             =   4320
      Width           =   975
   End
   Begin VB.Frame fraBrowser 
      Caption         =   "Web Form Browser"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   5760
      TabIndex        =   11
      Top             =   120
      Width           =   4335
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   3735
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   4095
         ExtentX         =   7223
         ExtentY         =   6588
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   3000
      Top             =   6360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdParse 
      Caption         =   "Parse"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Top             =   6360
      Width           =   975
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Open List"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   6360
      Width           =   975
   End
   Begin VB.Frame fraLinks 
      Caption         =   "Link Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   5535
      Begin VB.ListBox lstLinks 
         Height          =   1035
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraDescription 
      Caption         =   "Description Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   3
      Top             =   3240
      Width           =   5535
      Begin VB.ListBox lstDescription 
         Height          =   1035
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraTitle 
      Caption         =   "Title Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   5535
      Begin VB.ListBox lstTitle 
         Height          =   1035
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Frame fraSource 
      Caption         =   "Source"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.ListBox lstSource 
         Height          =   1035
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   5295
      End
   End
   Begin VB.Label lblAuthor 
      BackStyle       =   0  'Transparent
      Caption         =   "Author: Charles J Butler"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   6480
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##############################################
'#
'#  Strip Tags by Charles J Butler
'#  Email cbutler@defonic.com
'#
'##############################################

Private Sub Form_Load()
    
    '  Locate your form, we use a test form but this is intended for you web form.

    WebBrowser1.Navigate App.Path & "/form.htm"
    
End Sub

Private Sub cmdLoad_Click()
' clears any current data
lstSource.Clear
'sets common dialogs title to textbelow
    cd.DialogTitle = "Open Document"
    
'directory thats being show is the one the apps directorys at
    cd.InitDir = App.Path
    
'sets flag to &h4
    cd.Flags = &H4
    
'filters/allows certain file types
    cd.Filter = "All Files (*.*)|*.*"
    
'show the save button, and open it
    cd.ShowOpen
    
If cd.FileName = "" Then ' if user selects cancel do nothing

Else
    Call LoadList(cd.FileName, lstSource)
End If
fraSource.Caption = "Source (" & lstSource.ListCount & ") Total Lines Detected"
End Sub

Private Sub cmdParse_Click()
Dim x As Integer
Dim myString1, Mystring2, Mystring3 As String

    For x = 0 To lstSource.ListCount - 1 ' lets get the listcount

myString1 = GetTagText(lstSource.List(x), "title") ' lets get the text between <title>
Mystring2 = GetTagText(lstSource.List(x), "description") ' lets get the text between <description>
Mystring3 = GetTagText(lstSource.List(x), "link") ' lets get the text between <link>

' lets add our strings to the said locations
    lstTitle.AddItem myString1
    lstDescription.AddItem Mystring2
    lstLinks.AddItem Mystring3

Next


' just in case lets remove any duplicate entries if there are any
ListKillDuplicates lstTitle
ListKillDuplicates lstDescription
ListKillDuplicates lstLinks


fraTitle.Caption = "Title Tags (" & lstTitle.ListCount & ")"
fraDescription.Caption = "Description Tags (" & lstDescription.ListCount & ")"
fraLinks.Caption = "Link Tags (" & lstLinks.ListCount & ")"

End Sub


Private Sub cmdPost_Click()
On Error Resume Next
Dim x As Integer

For x = 0 To lstTitle.ListCount - 1

Pause 2 ' use a pause if connection is slow

WebBrowser1.Document.Forms(0).elements(0).Value = lstTitle.List(x) ' first textarea
WebBrowser1.Document.Forms(0).elements(1).Value = lstDescription.List(x) ' second text area
WebBrowser1.Document.Forms(0).elements(2).Value = lstLinks.List(x) ' third text area

Me.Caption = "Posting - " & lstTitle.List(x)
Pause 2 ' use a pause if connection is slow

WebBrowser1.Document.Forms("data").All("Submit").Click ' make sure you know your full form name when recreating this for your own


Next

lstSource.Clear
lstDescription.Clear
lstLinks.Clear

End Sub
