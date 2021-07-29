'//=================================================================//
'/|     MODULE:  basPatGenRecursionDemo                             |/
'/|    PURPOSE:  Generate Trim Patterns, honoring MiddleOnly        |/
'/|              Constraints and allowing for sizes to be trimmed   |/
'/|              at either the Cross Direction size or the Machine  |/
'/|              Direction                                          |/
'/|         BY:  Sean                                               |/
'/|       DATE:  12/07/15                                           |/
'//=================================================================//
Option Compare Database
Option Explicit


'// Variables
Private mvArr           As Variant
Private maResult()      As String
Private maInclude()     As Long
Private mlElementCount  As Long
Private mlResultCount   As Long


Sub AllCombos()
   'Call AllCombos
    Dim i As Long, z As Long

    'Initialize arrays and variables
    Erase maInclude
    Erase maResult
    mlResultCount = 0

    'Create array of possible substrings
    mvArr = Array("A", "B", "C", "D", "E", "F")
    ''(1, 2, 3, 4, 5, 6, 7, 8, 9)
    ''10, 11, 12, 13, 14, 15, 16, 17, 18, _
      ''           19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35)

    'Initialize variables based on size of array
    mlElementCount = UBound(mvArr)
    ReDim maInclude(LBound(mvArr) To UBound(mvArr))
    ReDim maResult(1 To 2 ^ (mlElementCount + 1))  'this is an overflow problem
    'ReDim maResult(1 To 50000)  '50K

    'Call the recursive function for the first time
    Eval 0

    'Print the results to the immediate window
    For i = LBound(maResult) To UBound(maResult)
        z = Len(maResult(i)) - Len(Replace(maResult(i), " ", "")) + 1
        If z >= 0 And z <= 9 Then
          Debug.Print i & Space(8 - Len(CStr(i))) & Trim(maResult(i)) & Space(20 - Len(Trim(CStr(maResult(i))))) & z
        End If
    Next i

End Sub


Sub Eval(ByVal lPosition As Long)

    Dim sConcat As String
    Dim i As Long

    If mlResultCount <= UBound(maResult) Then
    
       If lPosition <= mlElementCount Then
           'set the position to zero (don't include) and recurse
           maInclude(lPosition) = 0
           Eval lPosition + 1
   
           'set the position to one (include) and recurse
           maInclude(lPosition) = 1
           Eval lPosition + 1
       Else
           'once lPosition exceeds the number of elements in the array
           'concatenate all the substrings that have a corresponding 1
           'in maInclude and store in results array
           mlResultCount = mlResultCount + 1
           For i = 0 To UBound(maInclude)
               If maInclude(i) = 1 Then
                   sConcat = sConcat & mvArr(i) & Space(1)
               End If
           Next i
           If mlResultCount <= UBound(maResult) Then
               sConcat = Trim(sConcat)
               
               maResult(mlResultCount) = sConcat
           End If
       End If
       
    End If
   
End Sub
