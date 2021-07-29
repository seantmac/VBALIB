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
    Dim iSpacer As String

    'Initialize arrays and variables
    Erase maInclude
    Erase maResult
    mlResultCount = 0
    iSpacer = 88
   Debug.Print Now()
    'Create array of possible substrings
    mvArr = Array("A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L")
    ''mvArr = Array(99, 98, 97, 92, 88, 86, 84, 83, 81, 78, 20)
    
    '(300, 280, 275, 250, 240, 225, 200, 175, 160, 150, 140, 125, 80, 50, 25)
    '(212, 159, 110, 106, 98.5, 91, 53, 45, 44, 22)
    '("A", "B", "C", "D", "E", "F")
    '("A", "B", "C", "D", "E", "F", "G",  "H",  "I",  "J",  "K", "L")
    '(212, 159, 110, 106, 98.5, 91, 53, 45, 44, 22)
    ''(1, 2, 3, 4, 5, 6, 7, 8, 9)
    ''10, 11, 12, 13, 14, 15, 16, 17, 18, _
      ''           19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35)
      
    ''(107.75, 104.75, 103.85, 100.75, 98.875, 98.725, 97.875, 96.875, 95.875, 94.875, 92.875, 91.875, 91.725, 90.875, 87.875, 85.875, 84.875, 84.125, 83.875, 82.875, 80.875, 77.875, 55.275, 44.275, 33.275, 20.011)

    ''11 sizes
    ''(99, 98, 97, 92, 88, 86, 84, 83, 81, 78, 20)
      
    'Initialize variables based on size of array
    mlElementCount = UBound(mvArr)
    ReDim maInclude(LBound(mvArr) To UBound(mvArr))
    ''z = 2 ^ (mlElementCount + 1)
    z = (uMAX(1, (-2 * 5 + 1 * mlElementCount)) + mlElementCount ^ 5) / 0.88
    z = CLng(z)
    ReDim maResult(1 To z)  'this is an overflow problem
    'ReDim maResult(1 To 50000)  '50K

    'Call the recursive function for the first time
    Eval 0

   Debug.Print Now()

    'Print the results to the immediate window
    For i = LBound(maResult) To UBound(maResult)
        z = Len(maResult(i)) - Len(Replace(maResult(i), " ", "")) + 1
        If z >= 0 And z <= 9 Then
          If Len(Trim(maResult(i))) >= 1 Then Debug.Print i & Space(18 - Len(CStr(i))) & Trim(maResult(i)) & Space(iSpacer - Len(Trim(CStr(maResult(i))))) & z
        End If
    Next i

   Debug.Print Now()
   Debug.Print "Done."
End Sub


Sub Eval(ByVal lPosition As Long)

    Dim sConcat As String
    Dim i As Long
    sConcat = ""
   
    If mlResultCount <= UBound(maResult) Then
    
       '''''Debug.Print Space(5) & sConcat
   
       If lPosition <= mlElementCount Then
           'set the position to zero (don't include) and recurse
           maInclude(lPosition) = 0
           Eval lPosition + 1
   
           'set the position to one (include) and recurse
           maInclude(lPosition) = 1
           Eval lPosition + 1
           
           '''''Debug.Print Space(27) & sConcat & "  this was a non-finished sConcat"
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
           
           ''Debug.Print Space(25) & "**" & sConcat

       End If
       
    End If
   
End Sub






Function Factorial(N As Long) As Long
'//=========================================================================//
'/|   FUNCTION: Factorial                                                   |/
'/|      USAGE: i = Factorial(4)                                            |/
'/|       DATE: 12/16/15                                                    |/
'//=========================================================================//
'http://www.cpearson.com/excel/recursiveprogramming.aspx

    If N = 1 Then
        Factorial = 1
    Else
        Factorial = N * Factorial(N - 1)
        
    End If
End Function