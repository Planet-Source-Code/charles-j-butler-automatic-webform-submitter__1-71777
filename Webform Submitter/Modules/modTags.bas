Attribute VB_Name = "modTags"
'##############################################
'#
'#  Tag Module by Charles J Butler
'#  Email cbutler@defonic.com
'#
'##############################################

Public Function GetTag(strSource As String, Tag As String) As String

    'Returns full tag and text


    If InStr(strSource, "<" & Tag & ">") = 0 Then
        GetTag = ""
        Exit Function
    End If

    GetTag = Mid$(strSource, InStr(strSource, "<" & Tag & ">"), InStr(strSource, "</" & Tag & ">") + Len("</" & Tag & ">") - 1)
End Function



Public Function GetTagText(strSource As String, Tag As String) As String

    'Returns text between tags


    If InStr(strSource, "<" & Tag & ">") = 0 Then
        GetTagText = ""
        Exit Function
    End If

    GetTagText = Mid$(strSource, InStr(strSource, "<" & Tag & ">") + Len("<" & Tag & ">"), (InStr(strSource, "</" & Tag & ">")) - (InStr(strSource, "<" & Tag & ">") + Len("<" & Tag & ">")))
End Function

Public Function RemoveTag(strSource As String, Tag As String) As String

    'Removes tag from text

    If InStr(strSource, "<" & Tag & ">") = 0 Then
        RemoveTag = ""
        Exit Function
    End If

    RemoveTag = Left$(strSource, InStr(strSource, "<" & Tag & ">") - 1) & Mid$(strSource, InStrRev(strSource, "</" & Tag & ">") + Len("</" & Tag & ">"))
End Function

Public Sub RandomizeList(List As listbox)
Dim x As Integer, y As Integer, _
intNumberOfEntries As Integer, intRandomNumber As Integer, _
tmpText As String, tmpText2 As String

Randomize
    intNumberOfEntries = List.ListCount 'how many entries
        
    For y = 1 To 2 'random passes
        For x = 0 To intNumberOfEntries - 1 'go through each entry sequentially
            tmpText = List.List(x)                               'get sequential entry
            intRandomNumber = Int(intNumberOfEntries * Rnd)   'get a random entry
            tmpText2 = List.List(intRandomNumber)            '      "
            List.List(x) = tmpText2                              ' swap both
            List.List(intRandomNumber) = tmpText             '      "
        Next x
        
    Next y
End Sub


Public Sub ListKillDuplicates(listbox As listbox)
    Dim FirstCount As Long, SecondCount As Long
    On Error Resume Next
    For FirstCount& = 0& To listbox.ListCount - 1
        For SecondCount& = 0& To listbox.ListCount - 1
        DoEvents
            If LCase(listbox.List(FirstCount&)) Like LCase(listbox.List(SecondCount&)) And FirstCount& <> SecondCount& Then
                listbox.RemoveItem SecondCount&
                
            End If
        Next SecondCount&
    Next FirstCount&
End Sub
Public Sub LoadList(sLocation As String, lstListBox As listbox)
On Error GoTo dlgerror
Dim sCurrent As String
Dim i As Integer
lstListBox.Clear
Open sLocation For Input As #1
i = 0
Do Until EOF(1)
Line Input #1, sCurrent
lstListBox.AddItem sCurrent, i
i = i + 1
        
Loop
Close #1
Exit Sub
dlgerror:
MsgBox "An error has occured " & Err.Description
Exit Sub
End Sub

Public Sub Pause(interval)
    Dim current
        current = Timer
        Do While Timer - current < Val(interval)
            DoEvents
        Loop
End Sub
