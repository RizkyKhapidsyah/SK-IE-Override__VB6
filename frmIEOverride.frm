VERSION 5.00
Begin VB.Form frmIEOverride 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "IE Override"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4860
   Icon            =   "frmIEOverride.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   4785
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   60
         TabIndex        =   3
         Text            =   "Type location URL"
         Top             =   675
         Width           =   4155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Go"
         Enabled         =   0   'False
         Height          =   315
         Left            =   4275
         TabIndex        =   2
         Top             =   675
         Width           =   420
      End
      Begin VB.ListBox lstLinks 
         Height          =   840
         Left            =   60
         TabIndex        =   1
         Top             =   1620
         Width           =   4155
      End
      Begin VB.Label Label1 
         Caption         =   "This example shows how to override default popup menu on Internet Explorer for a given link:"
         Height          =   450
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   150
         Width           =   4560
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Double-click on iten that you want to override default popup menu:"
         Height          =   390
         Index           =   1
         Left            =   90
         TabIndex        =   4
         Top             =   1140
         Width           =   4560
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Menu mnuX 
      Caption         =   "x"
      Visible         =   0   'False
      Begin VB.Menu mnuSecondary 
         Caption         =   "Menu Hello"
         Index           =   0
      End
      Begin VB.Menu mnuSecondary 
         Caption         =   "Menu Got it"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmIEOverride"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' On Project, References menu:
' Set a reference to
' Microsoft Internet explorer (shdocvw.dll)
' and mshtml.dll library
Private WithEvents IE As SHDocVw.InternetExplorer
Attribute IE.VB_VarHelpID = -1
Private HDoc As HTMLDocument
Private WithEvents hAnchor As HTMLAnchorElement
Attribute hAnchor.VB_VarHelpID = -1
Private hElement As Object
Private strUrls As New Collection 'HTMLAnchorElement

Private Sub Command1_Click()
Screen.MousePointer = vbHourglass
If Text1.Text <> "" Then
    ' Creates IExplorer New instance
    Set IE = New SHDocVw.InternetExplorer
    With IE
        .Navigate Text1.Text    ' Navigates to url wrote on textbox
        .Visible = True         ' Shows Internet Explorer window
    End With
End If
' "hide" main window app.
'Me.WindowState = vbMinimized
Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next                    ' Clean object references
If Not IE Is Nothing Then
    IE.Quit
    Set hAnchor = Nothing
    Set IE = Nothing
End If
End Sub


Private Function hAnchor_oncontextmenu() As Boolean
' Override default popup menu and put our menu instead.
PopupMenu mnuX
End Function


Private Sub IE_DocumentComplete(ByVal pDisp As Object, URL As Variant)
On Error Resume Next
Screen.MousePointer = vbHourglass
If pDisp Is IE Then                     ' If all document is complete downloaded
    Set HDoc = IE.Document
    Me.Caption = IE.LocationURL
    lstLinks.Clear
    For Each hElement In HDoc.body.All
        Select Case hElement.tagName
'        Case "IMG"                      ' If HTML object is an image.
'            Dim pic As HTMLImg
'            Set pic = hElement
'            pic.Style.visibility = "hidden"
        Case "A"                        ' If HTML object is a link.
            ' Add links to list
            With lstLinks
                .AddItem hElement.innerText
                
                If .ListCount > 0 Then
'                    ReDim Preserve strUrls(.ListCount)
                    strUrls.Add hElement, CStr(.ListCount)
                End If
            End With
        End Select
    Next
End If
Screen.MousePointer = vbNormal
End Sub

Private Sub IE_OnQuit()
Unload Me
End Sub


Private Sub lstLinks_DblClick()
Set hAnchor = strUrls.Item(CStr(lstLinks.ListIndex + 1))
End Sub


Private Sub mnuSecondary_Click(Index As Integer)
Select Case Index                       ' Some stuff to show
    Case 0                              ' how it works.
        MsgBox "hello World!"
    Case 1
        MsgBox "Gotcha!"
End Select
End Sub

Private Sub Text1_Change()
Command1.Enabled = True
End Sub

Private Sub Text1_GotFocus()
With Text1
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub




