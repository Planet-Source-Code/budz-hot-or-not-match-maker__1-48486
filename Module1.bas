Attribute VB_Name = "Module1"

Sub hotornotvoter()
'2003 VB6 hot or not match maker - www.budzsoft.com
'personal page             - www.budzsack.com
'webmaster resources       - www.getmoneyquick.info
'use at your own risk

Dim duh As Variant
Form1.Text1.Text = "http://wwww.hotornot.com/m/?action=vote"
Do
On Error Resume Next
Form1.Text2 = Form1.Inet1.OpenURL("http://wwww.hotornot.com/m/?action=vote", icString)

Loop Until Not Form1.Inet1.StillExecuting



Call SaveText(Form1.Text2, "c:\hotornotbuffer.txt")
Form1.Text2 = ""
Form1.Text1 = ""
Call Striphot("c:\hotornotbuffer.txt")
On Error Resume Next
Kill ("c:\hotornotbuffer.txt")




End Sub
Public Function ExtractText(ByVal StringToSearch As String, ByVal StartingString As String, ByVal EndingString As String) As String
  Dim intStartOffset As Integer
  Dim intEndOffset As Integer
  
   ExtractText = ""
  intStartOffset = InStr(1, StringToSearch, StartingString)
  If intStartOffset > 0 Then
     ' The starting string is found, get everything after it.
     ' Add the length of the Starting String to exclude the Starting String
     StringToSearch = Mid$(StringToSearch, intStartOffset + Len(StartingString))
     
     intEndOffset = InStr(1, StringToSearch, EndingString)
     If intEndOffset > 0 Then
        ' The ending string was found, return everything before it.
        ' Subtract 1 to exclude the Ending String
        ExtractText = Left$(StringToSearch, intEndOffset - 1)
     Else
        ' The ending string was not found
        ' Return nothing, because there is no string between
        ' a start and end, unless both exist.
        ExtractText = vbNullString
     End If
  Else
     ' The starting string was not found
     ' Return nothing, because there is no string between
     ' a start and end, unless both exist.
     ExtractText = vbNullString
  End If

End Function
Public Sub timeout(duration As Long)

    
    Dim current As Long
    current = Timer
    Do Until Timer - current >= duration
        DoEvents
   
    Loop
End Sub
Public Function replacestring(strstring As String, strwhat As String, strwith As String) As String

    
    Dim lngpos As Long
    Do While InStr(1&, strstring$, strwhat$)
        DoEvents
        Let lngpos& = InStr(1&, strstring$, strwhat$)
        Let strstring$ = Left$(strstring$, (lngpos& - 1&)) & Right$(strstring$, Len(strstring$) - (lngpos& + Len(strwhat$) - 1&))
    Loop
    Let replacestring$ = strstring$
End Function
Public Function removechar(thestring As String, char As String) As String


    removechar$ = replacestring(thestring$, char$, "")
End Function
Function randomnum(Last)
Dim X As Long

X = Int(Rnd * Last + 1)
        randomnum = X
End Function
Public Sub savelistbox(strnameandpath As String, thelist As ListBox)

    Dim index As Long
    On Error Resume Next
    Open strnameandpath$ For Output As #1&
    For index& = 0& To thelist.ListCount - 1&
        Print #1&, thelist.List(index&)
    Next index&
    Close #1&
End Sub
Public Function randomletter() As String

    
    Dim Random As Long
    Randomize
    Random& = Int(Rnd * 26) + 1
    If Random& = 1 Then randomletter$ = "a"
    If Random& = 2 Then randomletter$ = "b"
    If Random& = 3 Then randomletter$ = "c"
    If Random& = 4 Then randomletter$ = "d"
    If Random& = 5 Then randomletter$ = "e"
    If Random& = 6 Then randomletter$ = "f"
    If Random& = 7 Then randomletter$ = "g"
    If Random& = 8 Then randomletter$ = "h"
    If Random& = 9 Then randomletter$ = "i"
    If Random& = 10 Then randomletter$ = "j"
    If Random& = 11 Then randomletter$ = "k"
    If Random& = 12 Then randomletter$ = "l"
    If Random& = 13 Then randomletter$ = "m"
    If Random& = 14 Then randomletter$ = "n"
    If Random& = 15 Then randomletter$ = "o"
    If Random& = 16 Then randomletter$ = "p"
    If Random& = 17 Then randomletter$ = "q"
    If Random& = 18 Then randomletter$ = "r"
    If Random& = 19 Then randomletter$ = "s"
    If Random& = 20 Then randomletter$ = "t"
    If Random& = 21 Then randomletter$ = "u"
    If Random& = 22 Then randomletter$ = "v"
    If Random& = 23 Then randomletter$ = "w"
    If Random& = 24 Then randomletter$ = "x"
    If Random& = 25 Then randomletter$ = "y"
    If Random& = 26 Then randomletter$ = "z"
End Function

Public Sub Striphot(FilePath As String)
Dim tmpEmail1, tmpEmail2 As String

Open FilePath For Input As #1
Do Until EOF(1)
Input #1, tmpEmail1
For X = 1 To Len(tmpEmail1)
    tmpEmail2 = Mid(tmpEmail1, X, 7)
    If tmpEmail2 = "http://meetme.hotornot.com/?" Then
        St1 = X
        tmpY = X + 1
        For Y = 1 To Len(tmpEmail1)
            tmpEmail2 = Mid(tmpEmail1, tmpY, 1)
            If tmpEmail2 = Chr(34) Then
                St2 = tmpY
                tmpEmail2 = Mid(tmpEmail1, St1 + 7, ((St2 - St1) - 7))
                If (Left(tmpEmail2, 2) <> "//") And (Left(tmpEmail2, 1) <> " ") Then
                    Form1.hon.Text = tmpEmail2
                    Exit For
                End If
            End If
            tmpY = tmpY + 1
        Next Y
    End If
Next X
Loop
Close #1
End Sub

Sub SaveText(txtSave As TextBox, Path As String)
    Dim TextString As String
    On Error Resume Next
    TextString$ = txtSave.Text
    Open Path$ For Output As #1
    Print #1, TextString$
    Close #1
End Sub

