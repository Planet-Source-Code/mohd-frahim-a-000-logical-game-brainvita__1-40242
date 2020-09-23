Attribute VB_Name = "Common"
Public highscore(10) As String
Public hightime(10) As String
Public highname(10) As String
Public highdate(10) As String

Public elapsedtime As Long

Public counter As Integer
Public previous As Integer
Public current As Integer
Public row As Integer
Public col As Integer
Public dstrow As Integer
Public dstcol As Integer
Public srcrow As Integer
Public srccol As Integer
Public adjrow(5) As Integer
Public adjcol(5) As Integer
Public moverow(5) As Integer
Public movecol(5) As Integer
Public score As Integer
Public backcount As Integer
Public fromboxrow(32) As Integer
Public toboxrow(32) As Integer
Public fromboxcol(32) As Integer
Public toboxcol(32) As Integer

Public paused As Boolean
Public movevalid(5) As Boolean
Public movespossible As Boolean
Public gameend As Boolean
Public selected As Boolean
Public boxfilled(9, 9) As Boolean
Public firsthigh As Boolean
Public highshow As Boolean

Public Function checkend()
Dim temp As Integer
gameend = False

For i = 2 To 4
    pictrowcol (i)
    If boxfilled(row, col) = True Then
        Call showmoves(row, col)
        If movespossible = True Then Exit Function
    End If
Next i

For i = 9 To 11
    pictrowcol (i)
    If boxfilled(row, col) = True Then
        Call showmoves(row, col)
        If movespossible = True Then Exit Function
    End If
Next i

For i = 14 To 34
    pictrowcol (i)
    If boxfilled(row, col) = True Then
        Call showmoves(row, col)
        If movespossible = True Then Exit Function
    End If
Next i

For i = 37 To 39
    pictrowcol (i)
    If boxfilled(row, col) = True Then
        Call showmoves(row, col)
        If movespossible = True Then Exit Function
    End If
Next i

For i = 44 To 46
    pictrowcol (i)
    If boxfilled(row, col) = True Then
        Call showmoves(row, col)
        If movespossible = True Then Exit Function
    End If
Next i

main.Timer1.Enabled = False
paused = False
main.cmdpause.Enabled = False
gameend = True

MsgBox "GAME OVER." & vbCrLf & "Your score is " & score & " out of 32" _
    & vbCrLf & endmsg, vbOKOnly + vbExclamation, "Brainvita"
chkhighscore

End Function
Public Function endmsg() As String
If score = 1 Then endmsg = "You are a GENIUS."
If score = 2 Or score = 3 Then endmsg = "You are OUTSTANDING."
If score = 4 Or score = 5 Then endmsg = "You are SUPERB."
If score >= 6 And score <= 8 Then endmsg = "You are GREAT."
If score >= 9 And score <= 10 Then endmsg = "You are GOOD."
If score >= 11 And score <= 15 Then endmsg = "You are OKAY.Practice more."
If score > 15 Then endmsg = "You need a lot of PRACTICE.Keep playing till you improve."
End Function

Public Function storeback()
Dim temp As Integer

If backcount > 0 Then main.cmdback.Enabled = True
If backcount < 31 Then
    backcount = backcount + 1
    temp = backcount
Else
    For i = 0 To 31
        fromboxrow(i) = fromboxrow(i + 1)
        fromboxcol(i) = fromboxcol(i + 1)
        toboxrow(i) = toboxrow(i + 1)
        toboxcol(i) = toboxcol(i + 1)
    Next i
    temp = 31
End If

pictrowcol (previous)
fromboxrow(temp) = row
fromboxcol(temp) = col
pictrowcol (current)
toboxrow(temp) = row
toboxcol(temp) = col

End Function
Public Function chkhighscore()
Dim tempstring As String
firsthigh = False
'On Error GoTo higherr
counter = 0
Open "BVHigh.brh" For Input As #1
Line Input #1, tempstring
If tempstring = "Start" Then GoTo skip1

Line Input #1, tempstring
highscore(counter) = tempstring
Line Input #1, tempstring
If EOF(1) Then GoTo skip1
hightime(counter) = tempstring
Line Input #1, tempstring
If EOF(1) Then GoTo skip1
highname(counter) = tempstring
Line Input #1, tempstring
'If EOF(1) Then GoTo skip1

GoTo skip2:

skip1:

If EOF(1) Then
    MsgBox "It seems you are playing this game for the first time." & vbCrLf & _
          "Your highscore will now be registered.", vbOKOnly + vbExclamation, "Brainvita"
    Close #1
    Open "BVHigh.brh" For Output As #1
    Print #1, "Highscores"
    Close #1
    firsthigh = True
    high.Show
    firsthigh = True
    Exit Function
End If

skip2:

While Not EOF(1)
    highdate(counter) = tempstring
    If score < highscore(counter) Then ' if score is among highest
        Close #1
        high.Show
        Close #1
        Exit Function
    Else  ' if score equal then compare time
        If score = highscore(counter) And elapsedtime <= hightime(counter) Then
            Close #1
            high.Show
            Exit Function
        End If
    End If
    counter = counter + 1
    Line Input #1, tempstring
    highscore(counter) = tempstring
    Line Input #1, tempstring
    hightime(counter) = tempstring
    Line Input #1, tempstring
    highname(counter) = tempstring
    Line Input #1, tempstring

Wend

If counter < 10 Then
    Close #1
    high.Show
    Exit Function
End If

Close #1
Exit Function

End Function

Public Function pictrowcol(pictindex As Integer)
'this function calculates the row and column from picture index
row = (pictindex \ 7)
col = pictindex - ((pictindex \ 7) * 7)
row = row + 1
col = col + 1
End Function
Public Function rowcolpict(boxrow As Integer, boxcol As Integer) As Integer
'this function calculates the picture index from rows and column
rowcolpict = ((boxrow - 1) * 7) + boxcol - 1
End Function

Public Function checkmove(dstrow As Integer, dstcol As Integer) As Boolean

For i = 1 To 4
    If moverow(i) = dstrow And movecol(i) = dstcol Then
        checkmove = True
        Exit Function
    End If
Next i

End Function
Public Function showmoves(selrow As Integer, selcol As Integer)
'This function show the possible moves from any box
'There can be four possible moves at max from any box

movespossible = True

'Four adjecent boxes' row and column will be

adjrow(1) = selrow + 1
adjcol(1) = selcol

adjrow(2) = selrow - 1
adjcol(2) = selcol

adjrow(3) = selrow
adjcol(3) = selcol + 1

adjrow(4) = selrow
adjcol(4) = selcol - 1



For i = 1 To 4
    movevalid(i) = True
    'Column and row should not be greater than 9 and less than 1
    If adjrow(i) > 7 Or adjrow(i) < 1 Or adjcol(i) > 7 Or adjcol(i) < 1 Then
        movevalid(i) = False
    Else
    'There should be ball on the boxes adjecent to current selected box
        If boxfilled(adjrow(i), adjcol(i)) = False Then
            movevalid(i) = False
        Else   ' if ball is there on adjecent to current selected
               ' calculate move
            Select Case i
            Case 1: moverow(i) = selrow + 2
                    movecol(i) = selcol
            Case 2: moverow(i) = selrow - 2
                    movecol(i) = selcol
            Case 3: moverow(i) = selrow
                    movecol(i) = selcol + 2
            Case 4: moverow(i) = selrow
                    movecol(i) = selcol - 2
            End Select
                'column and row should not be greater than 9 and less than 1
            If adjrow(i) > 7 Or adjrow(i) < 1 Or adjcol(i) > 7 Or adjcol(i) < 1 Then
                movevalid(i) = False
            Else  ' there should not be ball on the possible move
                If boxfilled(moverow(i), movecol(i)) = True Then movevalid(i) = False
            End If
        End If
            
    End If
Next i

If movevalid(1) = False And movevalid(2) = False And _
   movevalid(3) = False And movevalid(4) = False Then movespossible = False


End Function
Public Function updboard(upprev As Integer)
Dim tmprowp As Integer
Dim tmpcolp As Integer
Dim tmprowc As Integer
Dim tmpcolc As Integer

score = score - 1
pictrowcol (upprev)
tmprowp = row
tmpcolp = col
boxfilled(row, col) = False
main.board(upprev).Picture = main.emptypict.Picture

pictrowcol (current)
tmprowc = row
tmpcolc = col
boxfilled(row, col) = True
main.board(current).Picture = main.mainpict.Picture
If tmprowc = tmprowp Then
    boxfilled(tmprowc, (tmpcolc + tmpcolp) / 2) = False
    main.board(rowcolpict(tmprowc, (tmpcolc + tmpcolp) / 2)).Picture = main.emptypict.Picture
End If

If tmpcolc = tmpcolp Then
    boxfilled((tmprowp + tmprowc) / 2, tmpcolc) = False
    main.board(rowcolpict((tmprowp + tmprowc) / 2, tmpcolc)).Picture = main.emptypict.Picture
End If
main.lblscore.Caption = score & " remaining out of 32"
storeback
checkend
End Function

Public Function newgame()
main.cmdback.Enabled = False
main.Caption = "Brainvita"
For i = 0 To 48
    main.board(i).Enabled = True
Next i
    
gameend = False
score = 32
elapsedtime = 0
paused = False
main.cmdpause.Caption = "Pause"
main.Timer1.Enabled = True
main.cmdback.Enabled = False
backcount = 0
main.lbltime.Caption = "0 seconds"
main.lblscore.Caption = "32 remaining out of 32"
main.lblstatus = "Please select a box."
selected = False
previous = 100
main.cmdpause.Enabled = True
For i = 0 To 8
    For j = 0 To 8
        boxfilled(i, j) = True
    Next j
Next i

boxfilled(4, 4) = False

For i = 0 To 48
    main.board(i).Appearance = 1
    main.board(i).Picture = main.mainpict.Picture
    main.board(i).BackColor = &H8000000F
Next i

main.board(24).Picture = main.emptypict.Picture

End Function
