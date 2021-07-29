'//=================================================================//
'/|     MODULE:  basPatGen8                                         |/
'/|    PURPOSE:  Generate Trim Patterns, honoring MiddleOnly        |/
'/|              Constraints and allowing for sizes to be trimmed   |/
'/|              at either the Cross Direction size or the Machine  |/
'/|              Direction                                          |/
'/|         BY:  Sean                                               |/
'/|       DATE:  12/16/15                                           |/
'//=================================================================//
Option Compare Database
Option Explicit


'// Variables
Private mlLastErr       As Long
Private msLastErr       As String
Private msErrSource     As String
Private msErrDesc       As String


'Private mvArr           As Variant  'varrElement()
'Private maResult()      As String	 'tarrPattern()
'Private maInclude()     As Long	 'larrInclude()  '0 if no, 1 if yes
'Private mlElementCount  As Long	 'lElementCount
'Private mlResultCount   As Long	 'lResultCount


'// Constants
Global Const LPM_BIG_M = 999999999


'// User Types
Type typItem   'items to be put together into various combinations to create trim Patterns
   lngItemID           As Long
   strItemCode         As String
   strItemName         As String
   dblItemSize         As Double
   blnItemMiddleOnlyCnst As Boolean
End Type

Type typPattern   'Final Trim Pattern
   lngPatternID        As Long
   strPatternCode      As String
   strPatternName      As String
   dblPatternSize      As Double
   dblPatternWaste     As Double
   lngPatternItemCount As Long
   lngPatternComboID() As Long
End Type




Sub Eval(ByVal lPosition As Long)
   Dim sConcat  As String
   dIM sPattern as typPattern
   Dim i As Long

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
      sConcat = Trim(sConcat)
      maResult(mlResultCount) = sConcat
   End If
End Sub