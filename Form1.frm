VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                 Hot Or Not Match Maker 2k3 - budzsoft.com"
   ClientHeight    =   6420
   ClientLeft      =   4620
   ClientTop       =   3750
   ClientWidth     =   6585
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6585
   Begin VB.Frame Frame3 
      Caption         =   "user control interface"
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   5640
      Width           =   6135
   End
   Begin VB.CommandButton Command8 
      Caption         =   "pause"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   6120
      Width           =   735
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   5640
      TabIndex        =   20
      Text            =   "0"
      Top             =   6120
      Width           =   735
   End
   Begin VB.ListBox List1 
      Height          =   3375
      Left            =   240
      TabIndex        =   19
      Top             =   6480
      Width           =   2775
   End
   Begin VB.Frame Frame2 
      Caption         =   "login please"
      Height          =   255
      Left            =   240
      TabIndex        =   18
      Top             =   0
      Width           =   6135
   End
   Begin VB.Frame Frame1 
      Caption         =   "persons keywords"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   4560
      Width           =   6135
   End
   Begin VB.TextBox personkw 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   4800
      Width           =   6135
   End
   Begin VB.ComboBox keyword 
      Height          =   315
      Left            =   3720
      TabIndex        =   15
      Text            =   "{keyword}"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   525
      Left            =   6720
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4800
      Width           =   6495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   4455
      Left            =   240
      TabIndex        =   13
      Top             =   240
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   7858
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
      Location        =   "http:///"
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   2520
      TabIndex        =   12
      Text            =   "of any age"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1320
      TabIndex        =   11
      Text            =   "women"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Text            =   "straight"
      Top             =   5880
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "login"
      Height          =   255
      Left            =   4920
      TabIndex        =   9
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox picurl 
      Height          =   495
      Left            =   7680
      TabIndex        =   8
      Text            =   "Text3"
      Top             =   10200
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   4335
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   6135
      ExtentX         =   10821
      ExtentY         =   7646
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
      Location        =   "http:///"
   End
   Begin VB.TextBox hon 
      Height          =   285
      Left            =   7560
      TabIndex        =   6
      Text            =   "Text3"
      Top             =   9720
      Width           =   1215
   End
   Begin VB.TextBox rand 
      Height          =   285
      Left            =   7560
      TabIndex        =   5
      Text            =   "Text3"
      Top             =   9360
      Width           =   1215
   End
   Begin VB.TextBox vt 
      Height          =   285
      Left            =   7560
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   9000
      Width           =   1215
   End
   Begin VB.TextBox votee 
      Height          =   285
      Left            =   7560
      TabIndex        =   3
      Text            =   "Text3"
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   255
      Left            =   5640
      TabIndex        =   2
      Top             =   5880
      Width           =   735
   End
   Begin VB.TextBox Text2 
      Height          =   2055
      Left            =   8400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   5520
      Width           =   4695
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   7440
      Top             =   7560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   8400
      TabIndex        =   0
      Top             =   7680
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()



If Frame2.Caption = "login please" Then MsgBox "login first please": GoTo done

Frame2.Caption = "starting up ;]"
Do
timeout 0.04
WebBrowser2.Left = WebBrowser2.Left + 20
Loop Until WebBrowser2.Left >= 6720

more:

timeout Val(Combo4.Text)
'Call hotornotvoter
If Text4.Text = "" Then
WebBrowser2.Navigate ("http://wwww.hotornot.com/m/?action=vote")
Do
timeout 0.04
Loop Until WebBrowser2.Busy = False

Text2.Text = WebBrowser2.Document.documentelement.innerhtml

GoTo wooa
End If
timeout 0.04
WebBrowser2.Navigate (Text4.Text)

Do
timeout 0.04
Loop Until WebBrowser2.Busy = False
timeout 1
Text2.Text = ""
Text2.Text = WebBrowser2.Document.documentelement.innerhtml
wooa:

Do
timeout 0.04
Loop Until WebBrowser2.Busy = False
Text2.Text = WebBrowser2.Document.documentelement.innerhtml
votee.Text = ExtractText(Text2, "votee=", "&")
vt.Text = ExtractText(Text2, "vt=", "&")
rand.Text = ExtractText(Text2, ";rand=", " target")
rand.Text = removechar(rand.Text, Chr(34))





hon.Text = ExtractText(Text2, "meetme.hotornot.com/?", " ")
hon.Text = removechar(hon.Text, Chr(34))


''''''

If InStr(Text2.Text, ".JPG") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".JPg")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".JPG"

GoTo dp
End If
If InStr(Text2.Text, ".jpg") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".jpg")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".jpg"

GoTo dp
End If

If InStr(Text2.Text, ".jpeg") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".jpeg")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".jpeg"
End If

If InStr(Text2.Text, ".GIF") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".GIF")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".GIF"
End If
If InStr(Text2.Text, ".gif") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".gif")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".gif"
End If
dp:
If picurl.Text = "http://pix2.hotornot.com/pics.JPG" Then GoTo npic
If picurl.Text = "http://pix2.hotornot.com/pics.GIF" Then GoTo npic
If picurl.Text = "http://pix2.hotornot.com/pics.jpg" Then GoTo npic
If picurl.Text = "http://pix2.hotornot.com/pics.gif" Then GoTo npic
WebBrowser1.Navigate (picurl.Text)
Do
timeout 0.05
Loop Until WebBrowser1.Busy = False
npic:
personkw.Text = ""

personkw.Text = ExtractText(Text2, "hints = '", "';")

If personkw.Text = "" Then

personkw.Text = personkw.Text & ExtractText(Text2, "My keywords:", "</td></table>")
End If
'''''

nav1 = "http://meetme.hotornot.com/?"
nav2 = "&state=vote&votee="
nav3 = "&vt="
nav4 = "&advMode=1&geoMode=1&rand="
nav5 = "&vote=yes"
nav6 = "&sel_age="
nav7 = "&sel_orient="
nav8 = "&sel_sex="

If Combo2.Text = "women" Then
ss = "female"
GoTo ng
End If

ss = "male"
ng:



If Combo1.Text = "straight" Then
so = "s"
GoTo dec
End If

so = "g"

dec:


If Combo3.Text = "of any age" Then
sa = "x"
End If
If Combo3.Text = "18-25" Then
sa = "1"
End If
If Combo3.Text = "26-32" Then
sa = "2"
End If
If Combo3.Text = "33-40" Then
sa = "3"
End If
If Combo3.Text = "40 +" Then
sa = "4"
End If



If keyword.Text = "{keyword}" Then
Text4.Text = nav1 & hon.Text & nav2 & votee.Text & nav3 & vt.Text & nav4 & rand.Text & nav5 & nav6 & sa & nav7 & so & nav8 & ss
 GoTo nav
End If

If keyword.Text = "" Then
Text4.Text = nav1 & hon.Text & nav2 & votee.Text & nav3 & vt.Text & nav4 & rand.Text & nav5 & nav6 & sa & nav7 & so & nav8 & ss
 GoTo nav
End If

Text4.Text = nav1 & hon.Text & nav2 & votee.Text & nav3 & vt.Text & nav4 & rand.Text & nav5 & "&keyword=" & keyword & nav6 & sa & nav7 & so & nav8 & ss

 


nav:






WebBrowser2.Navigate (Text4.Text)
If List1.ListCount > 150 Then
List1.Clear
End If

'List1.AddItem (picurl.Text & ":" & matchurl)
'to be able to flip thru or to previous ^ not finished


Frame2.Caption = "picture last clicked yes to"



timeout 0.04

GoTo more
done:

End Sub


Private Sub Command2_Click()
If InStr(Text2.Text, ".jpg") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".jpg")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".jpg"

GoTo dp
End If

If InStr(Text2.Text, ".jpg") Then
picurl.Text = ExtractText(Text2, "pix2.hotornot.com/pics", ".gif")
picurl.Text = "http://pix2.hotornot.com/pics" & picurl.Text & ".gif"
End If
dp:


End Sub

Private Sub Command3_Click()
WebBrowser2.Navigate ("http://meetme.hotornot.com/?d8ce=55502&state=vote&votee=231443&vt=&advMode=1&geoMode=1&rand=32&vote=yes")

Do
timeout 0.04
Loop Until WebBrowser2.Busy = False
Text2.Text = WebBrowser2.Document.documentelement.innerhtml
votee.Text = ExtractText(Text2, "votee=", "&")
vt.Text = ExtractText(Text2, "vt=", "&")
rand.Text = ExtractText(Text2, ";rand=", " target")
rand.Text = removechar(rand.Text, Chr(34))




hon.Text = ExtractText(Text2, "meetme.hotornot.com/?", " ")
hon.Text = removechar(hon.Text, Chr(34))


End Sub

Private Sub Command4_Click()

WebBrowser2.Navigate ("http://meetme.hotornot.com/")

Text2.Text = WebBrowser2.Document.documentelement.innerhtml




Frame2.Caption = "thank you for logging in ;]"
End Sub

Private Sub Command5_Click()

End Sub

Private Sub Command6_Click()
timeout 2
hon.Text = ExtractText(Text2, "meetme.hotornot.com/?", " ")
hon.Text = removechar(hon.Text, Chr(34))

votee.Text = ExtractText(Text2, "&votee=", "&vt")
vt.Text = ExtractText(Text2, "vt=", "&advMode")
rand.Text = ExtractText(Text2, "&rand=", " ")
rand.Text = removechar(rand.Text, Chr(34))

End Sub

Private Sub Command7_Click()

End Sub

Private Sub Command8_Click()


Frame2.Caption = "paused for " & Combo4.Text & " seconds"
timeout Val(Combo4.Text)



End Sub

Private Sub Form_Load()

Combo1.AddItem "straight"
Combo1.AddItem "gay"
Combo2.AddItem "women"
Combo2.AddItem "men"
Combo3.AddItem "of any age"
Combo3.AddItem "18-25"
Combo3.AddItem "26-32"
Combo3.AddItem "18-25"
Combo3.AddItem "33-40"
Combo3.AddItem "40 +"
Combo4.AddItem "0"
Combo4.AddItem "5"
Combo4.AddItem "10"
Combo4.AddItem "15"
Combo4.AddItem "20"
Combo4.AddItem "25"
Combo4.AddItem "30"
Combo4.AddItem "40"
Combo4.AddItem "50"
Combo4.AddItem "60"
keyword.AddItem "music"
keyword.AddItem "sex"
keyword.AddItem "freak"
keyword.AddItem "encounter"
keyword.AddItem "cyber"
keyword.AddItem "horny"
keyword.AddItem "single"
keyword.AddItem "relationship"
keyword.AddItem "passionate"
keyword.AddItem "romantic"
keyword.AddItem "oral"
keyword.AddItem "drugs"
keyword.AddItem "fun"

WebBrowser1.Navigate ("http://www.hotornot.com/?f525=62757&state=rate&ratee=2676385&rk=VYK&2676385=10&rate=10")
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Unload Me
End
End Sub

Private Sub Form_Terminate()
Unload Me
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End
End Sub

