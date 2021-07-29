Function GetScoreDisplay(sPlayer As String, iRound As Integer) As Variant

'PASSING IN iRound=7 meand tourney over, lookup in column 10 for TOTAL TOURNEY SCORE

Dim nrGolfScores As Range        'lkpGolfScores
Dim nrCurrentRound As Range      'lkpCurrentRound
Dim nrParForTheCourse As Range   'lkpParForTheCourse
Dim iCountErr13 As Integer

Dim IsToday As Integer
Dim sTiming As String
Dim vScore As Variant
Dim vThru As Variant
Const COLUMN_OF_TODAYSCORE = 8
Const COLUMN_OF_TODAYSTHRU = 9

iCountErr13 = 0

On Error GoTo GetScoreDisplay_ERR
If iRound = 0 Then
   vThru = 1
End If

Set nrGolfScores = Sheets("Leaderboard").Range("lkpGolfScores")
Set nrParForTheCourse = Sheets("Tadboard").Range("lkpParForTheCourse")
Set nrCurrentRound = Sheets("Tadboard").Range("lkpCurrentRound")

vScore = Application.WorksheetFunction.VLookup(sPlayer, nrGolfScores, iRound + 3, False)
vThru = Application.WorksheetFunction.VLookup(sPlayer, nrGolfScores, COLUMN_OF_TODAYSTHRU, False)

'Set sTiming
If nrCurrentRound.Value > iRound Then sTiming = "AFTER"  'Previous Round Score Requested
If nrCurrentRound.Value < iRound Then sTiming = "PRE"    'Round Not Started yet for Player
If nrCurrentRound.Value = iRound Then
   If vThru = "F" Then
      sTiming = "AFTER"                            'Today's Round Requested But Player is Done
   ElseIf 0.1 < vScore And vScore < 0.9 Then
      sTiming = "PRE"                              'Today's Round Requested But Player has yet to tee off
   Else
      sTiming = "DURING"                           'Today's Round Requested.  Player is mid-Round
   End If
End If

If iRound = 7 Then sTiming = "AFTER"               'Tourney is over, Final Total Requested


'==========
Select Case sTiming
Case "AFTER"                                       ' -,MC,WD,DQ,77,61,59, etc.
   If vScore = "-" Or vScore = "E" Then
      GetScoreDisplay = vScore
   ElseIf vScore = "MC" Or vScore = "WD" Or vScore = "DQ" Then 'Handle DQ's, etc.
      GetScoreDisplay = vScore
   Else
      If iRound = 7 Then
         GetScoreDisplay = vScore
      Else
         GetScoreDisplay = vScore - nrParForTheCourse
      End If
   End If
Case "PRE"                                         '-,MC,WD,DQ,0.3625,0.4,0.4125 etc.
   If vScore = "-" Then
      GetScoreDisplay = vScore
   ElseIf vScore = "MC" Or vScore = "WD" Or vScore = "DQ" Then     'Handle DQ's, etc.
      GetScoreDisplay = vScore
   ElseIf 0.1 < vScore < 0.9 Then
      If Len(VBA.Minute(vScore)) = 1 Then
         GetScoreDisplay = VBA.Hour(vScore) & ":0" & VBA.Minute(vScore)
      Else
         GetScoreDisplay = VBA.Hour(vScore) & ":" & VBA.Minute(vScore)
      End If
   End If
Case "DURING"
   'Switch to TodayColumn -,MC,WD,DQ,E,-3,+5, etc.  Note that MC,WD,DQ must be written over "-" in Leaderboard
   vScore = Application.WorksheetFunction.VLookup(sPlayer, nrGolfScores, COLUMN_OF_TODAYSCORE, False)
   If vScore = "-" Or vScore = "E" Then
      GetScoreDisplay = vScore
   ElseIf vScore = "MC" Or vScore = "WD" Or vScore = "DQ" Then           'Handle DQ's, etc.
      GetScoreDisplay = vScore
   Else
      GetScoreDisplay = vScore
   End If

Case Else
   MsgBox "=== GetScoreDisplay === Before, During, or After?  It is not clear."

End Select
'==========

Exit Function

GetScoreDisplay_ERR:
   Select Case Err
   Case 1004 'Missing Golfer
      GetScoreDisplay = ""
      Exit Function
   Case 91   'BAD VARIABLE USE
      GetScoreDisplay = "bad"
      Exit Function
   Case 13   'Type Mismatch
      If iCountErr13 = 0 Then
         iCountErr13 = iCountErr13 + 1
         Resume
      Else
         Exit Function
      End If
   Case Else
      'MsgBox "=== GetScoreDisplay === " & Err.Number & " - " & Err.Description & "  " & sPlayer
      GetScoreDisplay = "???"
   End Select
   'Exit Function
   Resume Next
   Resume

End Function