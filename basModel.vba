'//=================================================================//
'/|     MODULE:  basModel                                           |/
'/|    PURPOSE:  Utility Functions for Modeling, and Solving        |/
'/|         BY:  Sean                                               |/
'/|       DATE:  08/27/97                                           |/
'/|    HISTORY:  08/27/97    Initial Release                        |/
'//=================================================================//
Option Compare Database
Option Explicit

'// Variables
Dim MPSLP As typMPSLP
Dim gMatGenOK As Boolean
Dim gFormRunFlag As Boolean ' set to true if running from form
Dim gtxtDest As TextBox

'// Constants
Global Const DAT_TABLE_PRE As String = "u"
Global Const DAT_CLASS_TEXT_LENGTH = 6
Global Const DAT_MAX_CLASSES = 16
Global Const LPM_TEXT_LENGTH = 80
Global Const LPM_DOUBLE_DECIMAL = 6
Global Const LPM_MAX_ELEMENT_CLASSES = 16

Global Const LPM_ROW_ELEMENT_TABLE_PRE As String = "tMTXtbliRow"      '// Matrix Elements
Global Const LPM_COL_ELEMENT_TABLE_PRE As String = "tMTXtbliCol"

Global Const LPM_CONSTR_DEF_TABLE_NAME As String = "tMTXtblDefRow"
Global Const LPM_COLUMN_DEF_TABLE_NAME As String = "tMTXtblDefCol"
Global Const LPM_COEFFS_DEF_TABLE_NAME As String = "tMTXtblDefCoeff"
Global Const LPM_DAT_DEF_TABLE_NAME As String = "tMTXtblDefDat"
Global Const LPM_FLD_DEF_TABLE_NAME As String = "tMTXtblDefFld"

Global Const LPM_CONSTR_TABLE_NAME As String = "tMTXtblMtxRow"
Global Const LPM_COLUMN_TABLE_NAME As String = "tMTXtblMtxCol"
Global Const LPM_MATRIX_TABLE_NAME As String = "tMTXtblMtxMatrix"

Global Const LPM_CONSTR_TABLE_NAME_IMPORT As String = "tMTXtblMtxRowSlnImport"
Global Const LPM_COLUMN_TABLE_NAME_IMPORT As String = "tMTXtblMtxColSlnImport"

Global Const LPM_BIG_M = 999999999

Global Const QG_CLASS_IN_BOTH = 1
Global Const QG_CLASS_IN_COL_ONLY = 2
Global Const QG_CLASS_IN_ROW_ONLY = 3



'// User Types
Type typCoeffType
   lngID As Long
   intActive As Integer
   strType As String
   strColType As String
   lngColID As Long
   strRowType As String
   lngRowID As Long
   strRecSet As String
   strCoeffFld As String
End Type

Type typColType
   lngID As Long
   intActive As Integer
   strType As String
   strDesc As String
   strTable As String
   strRecSet As String
   strPrefix As String
   strSOS As String
   intFREE As Integer
   strOBJFld As String
   strLOFld As String
   strUPFld As String
   intClassCount As Integer
   strClasses() As String
End Type

Type typDatType
   lngID As Long
   intActive As Integer
   intMaster As Integer
   strType As String
   strDesc As String
   strTable As String
   intClassCount As Integer
   strClasses() As String
   intFieldCount As Integer
End Type

Type typMPSLP
   strFileName As String
   strProblemName As String
   intOBJSense As Integer
   strOBJRowName As String
   strRowRS As String
   strColRS As String
   strCoeRS As String
End Type

Type typRowType
   lngID As Long
   intActive As Integer
   strType As String
   strDesc As String
   strTable As String
   strRecSet As String
   strPrefix As String
   strSense As String
   strRHSFld As String
   intClassCount As Integer
   strClasses() As String
End Type



Function BigM()
BigM = LPM_BIG_M
End Function

Function Brack(strIn As String) As String
   
   Brack = "[" & strIn & "]"

End Function



Function CountColClasses(strColType As String) As Integer
   
   Dim strColClassField As String, strColClass As String
   Dim i As Integer, intCount As Integer, intIgNull As Integer
   
   On Error GoTo CountColClasses_Err
      
   CountColClasses = -2
   intCount = 0
   intIgNull = False
      
   For i = 1 To LPM_MAX_ELEMENT_CLASSES      '// Count the Column classes
      strColClass = ""
      strColClassField = "C" & CStr(i)
      On Error Resume Next
      strColClass = DLookup(strColClassField, LPM_COLUMN_DEF_TABLE_NAME, "ColType = " & "'" & strColType & "'")
      If Len(strColClass) = 0 Then Exit For
      intCount = intCount + 1
   Next i
   
   CountColClasses = intCount

   
CountColClasses_Done:
  Exit Function

CountColClasses_Err:
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         CountColClasses = -2
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume CountColClasses_Done
      End If
   Case Else
      CountColClasses = -2
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CountColClasses_Done
      Resume
End Select

End Function

Function CountDatClasses(strDatType As String) As Integer
   
   Dim strDatClassField As String, strDatClass As String
   Dim i As Integer, intCount As Integer, intIgNull As Integer
   
   On Error GoTo CountDatClasses_Err
      
   CountDatClasses = -2
   intCount = 0
   intIgNull = False
      
   For i = 1 To DAT_MAX_CLASSES      '// Count the Dat classes
      strDatClass = ""
      strDatClassField = "D" & CStr(i)
      On Error Resume Next
      strDatClass = DLookup(strDatClassField, LPM_DAT_DEF_TABLE_NAME, "DatType = " & "'" & strDatType & "'")
      If Len(strDatClass) = 0 Then Exit For
      intCount = intCount + 1
   Next i
   
   CountDatClasses = intCount

   
CountDatClasses_Done:
  Exit Function

CountDatClasses_Err:
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         CountDatClasses = -2
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume CountDatClasses_Done
      End If
   Case Else
      CountDatClasses = -2
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CountDatClasses_Done
      Resume
End Select

End Function
Function CountDatFields(strDatType As String) As Integer
   
   Dim intCount As Integer, intIgNull As Integer
   Dim rs As Recordset, rsClone As Recordset, db As Database
   Dim strSQL As String
      
   On Error GoTo CountDatFields_Err
      
   CountDatFields = -2
   intCount = 0
   intIgNull = False
   
   strSQL = "SELECT DISTINCTROW " & LPM_DAT_DEF_TABLE_NAME & ".DatType, Count(" & LPM_DAT_DEF_TABLE_NAME & ".FldName) AS CountFields" & CRLF()
   strSQL = strSQL & "FROM " & LPM_DAT_DEF_TABLE_NAME & " " & CRLF()
   strSQL = strSQL & "GROUP BY " & LPM_DAT_DEF_TABLE_NAME & ".DatType" & CRLF()
   strSQL = strSQL & "HAVING (((" & LPM_DAT_DEF_TABLE_NAME & ".DatType)='" & strDatType & "'));"
   
   'Debug.Print strSQL
   
   Set db = CurrentDb
   Set rs = db.OpenRecordset(strSQL)
   Set rsClone = rs.Clone()
   rsClone.MoveFirst
   
   CountDatFields = rsClone!CountFields
   
CountDatFields_Done:
  Exit Function

CountDatFields_Err:
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         CountDatFields = -2
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume CountDatFields_Done
      End If
   Case 3021 'No records
      CountDatFields = 0
      Resume CountDatFields_Done
   Case Else
      CountDatFields = -2
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CountDatFields_Done
      Resume
End Select

End Function


Function CountRowClasses(strRowType As String) As Integer
   
   Dim strRowClassField As String, strRowClass As String
   Dim i As Integer, intCount As Integer, intIgNull As Integer
      
   On Error GoTo CountRowClasses_Err
      
   CountRowClasses = -2
   intCount = 0
      
   For i = 1 To LPM_MAX_ELEMENT_CLASSES      '// Count the Row Classes
      strRowClass = ""
      strRowClassField = "R" & CStr(i)
      On Error Resume Next
      strRowClass = DLookup(strRowClassField, LPM_CONSTR_DEF_TABLE_NAME, "RowType = " & "'" & strRowType & "'")
      If Len(strRowClass) = 0 Then Exit For
      intCount = intCount + 1
   Next i
   
   CountRowClasses = intCount

   
CountRowClasses_Done:
  Exit Function

CountRowClasses_Err:
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         CountRowClasses = -2
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume CountRowClasses_Done
      End If
   Case Else
      CountRowClasses = -2
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CountRowClasses_Done
      Resume
End Select

End Function

Function CreateColTable(colTemp As typColType)
'//================================================================================//
'/|   FUNCTION: CreateColTable                                                     |/
'/| PARAMETERS: colTemp, the column type to generate a table for.                  |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Creates a Table "zColPurLogMixPur" whose name is based on the      |/
'/|             standard prefix constant and the entry in the "ColTypeTable" field |/
'/|             of table LPM_COLUMN_DEF_TABLE_NAME.                                |/
'/|    PURPOSE: Create table for storing individual LP vectors                     |/
'/|      USAGE: i = CreateColTable(colNew)                                         |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/10/97                                                            |/
'/|    HISTORY: 4/2/97   Changed to use the user defined type for col types        |/
'//================================================================================//

Dim db As Database
Dim td As TableDef
Dim fd As Field
Dim ix As Index
Dim i As Integer
Dim lngID As Long

On Error GoTo CreateColTable_Err
 
   CreateColTable = False
   
   Set db = CurrentDb                                                         '// Return Current Database.
   Set td = db.CreateTableDef(LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable)   '// Create New Table Definition
   
   '// ColType ID Field
   Set fd = td.CreateField(colTemp.strType & "ID", dbLong)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = "0"
      'fd.Properties("Caption") = colTemp.strType & " ID"
   td.Fields.Append fd
   
   '// Add Primary Key Index
   Set ix = td.CreateIndex("PrimaryKey")
      ix.Primary = True
      Set fd = ix.CreateField(colTemp.strType & "ID", dbLong)
      ix.Fields.Append fd
   td.Indexes.Append ix
   
   '//  ColType Code Field
   Set fd = td.CreateField(colTemp.strType & "Code", dbText)
      fd.Size = LPM_TEXT_LENGTH
      'fd.Properties("Caption") = colTemp.strType & " Code"
   td.Fields.Append fd
   
   '//  Special Ordereed Sets (SOS) Field
   Set fd = td.CreateField("SOS", dbLong)
      fd.Properties("DefaultValue") = "0"
      fd.Properties("ValidationRule") = "0 Or 1 Or 2 Or 3"
      fd.Properties("ValidationText") = "0:  No Special Ordered Set for this Column Type. & CRLF() & 1:  At most one variable is non-zero." & CRLF() & "2:  At most 2 variables may be non-zero, and if two they must be adjacent." & CRLF() & "3:  Like Type 1, but all variables are 0/1." & CRLF() & "Please choose 1,2, or 3."
   td.Fields.Append fd
  
   '//  'Free Vector' Marker Field
   Set fd = td.CreateField("FREE", dbBoolean)
      fd.Properties("DefaultValue") = "No"
      'fd.Properties("Format") = "Yes/No"
   td.Fields.Append fd
   
   '// Add The Classes
   '// ----------------
   For i = 1 To colTemp.intClassCount
      Set fd = td.CreateField(colTemp.strClasses(i), dbText)
         fd.Size = LPM_TEXT_LENGTH
      td.Fields.Append fd
      
      '// Add Index
      Set ix = td.CreateIndex(colTemp.strClasses(i))
         Set fd = ix.CreateField(colTemp.strClasses(i), dbText)
         ix.Fields.Append fd
      td.Indexes.Append ix
   Next
   
   '// OBJ Field
   Set fd = td.CreateField("OBJ", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = LPM_DOUBLE_DECIMAL
   td.Fields.Append fd
   
   '// LO Field
   Set fd = td.CreateField("LO", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = "1"
   td.Fields.Append fd
   
   '// UP Field
   Set fd = td.CreateField("UP", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = "1"
   td.Fields.Append fd
      
   '// FX Field
   Set fd = td.CreateField("FX", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = "1"
   td.Fields.Append fd
   
   '// ACTIVITY Field
   Set fd = td.CreateField("ACTIVITY", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = LPM_DOUBLE_DECIMAL
   td.Fields.Append fd
      
   '// DJ VALUE Field
   Set fd = td.CreateField("DJ", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = LPM_DOUBLE_DECIMAL
   td.Fields.Append fd
   
   '// OBJ COEFF RANGE LO Field
   Set fd = td.CreateField("OBJLO", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = LPM_DOUBLE_DECIMAL
   td.Fields.Append fd
   
   '// OBJ COEFF RANGE UP Field
   Set fd = td.CreateField("OBJUP", dbDouble)
      'fd.Properties("Format") = "Fixed"
      'fd.Properties("DecimalPlaces") = LPM_DOUBLE_DECIMAL
   td.Fields.Append fd
   
   
   db.TableDefs.Append td
   db.TableDefs.Refresh
   
   CreateColTable = True

CreateColTable_Done:
  Exit Function

CreateColTable_Err:
Select Case Err
   Case 3010 'Table Already Exists
      DeleteTable (LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable)
      Resume
   Case 3211 'Table Currently in Use
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & " Please close the table (" & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & ") and hit OK to continue"
      Resume
   Case 3219 'Illegal Operation (Property Already Exists)
      Resume Next
   Case Else
      CreateColTable = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CreateColTable_Done
End Select

End Function
Function CreateDatTable(datTemp As typDatType)
'//================================================================================//
'/|   FUNCTION: CreateDatTable                                                     |/
'/| PARAMETERS: datTemp, the dat type to generate a table for.                     |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Creates a Table "zDatPurLogMixPur" whose name is based on the      |/
'/|             standard prefix constant and the entry in the "DatTypeTable" field |/
'/|             of table LPM_DAT_DEF_TABLE_NAME.                                                  |/
'/|    PURPOSE: Create table for storing individual Data Records                   |/
'/|      USAGE: i = CreateDatTable(datNew)                                         |/
'/|         BY: Sean                                                               |/
'/|       DATE: 4/4/97                                                             |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database, rs As Recordset, rsClone As Recordset, PR As Property
Dim td As TableDef
Dim fd As Field
Dim ix As Index
Dim i As Integer
Dim lngID As Long
Dim strSQL As String
Dim lngFieldDataType As Long

On Error GoTo CreateDatTable_Err
 
   CreateDatTable = False
   
   Set db = CurrentDb                                                         '// Return Current Database.
   Set td = db.CreateTableDef(DAT_TABLE_PRE & datTemp.strTable)               '// Create New Table Definition
   
   If datTemp.intMaster Or InStr(datTemp.strDesc, "Master") Then
      '// DatType ID Field
      Set fd = td.CreateField(datTemp.strType & "ID", dbLong)
      td.Fields.Append fd
      
      '// Add Primary Key Index
      Set ix = td.CreateIndex("PrimaryKey")
         ix.Primary = True
         Set fd = ix.CreateField(datTemp.strType & "ID", dbLong)
         ix.Fields.Append fd
      td.Indexes.Append ix
      
      '//  DatType Code Field
      Set fd = td.CreateField(datTemp.strType & "Code", dbText)
         fd.Size = LPM_TEXT_LENGTH
      td.Fields.Append fd
      
      '// Add Index for Field
      Set ix = td.CreateIndex(datTemp.strType & "Code")
      Set fd = ix.CreateField(datTemp.strType & "Code", dbText)
         ix.Fields.Append fd
      td.Indexes.Append ix
      
      '//  DatType Record Description Field
      Set fd = td.CreateField(datTemp.strType & "Desc", dbText)
         fd.Size = LPM_TEXT_LENGTH * 2
      td.Fields.Append fd
   End If
      
   '//  DatType Active
   Set fd = td.CreateField(datTemp.strType & "Active", dbBoolean)
      fd.Properties("DefaultValue") = "Yes"
      'fd.Properties("Format") = "Yes/No"
      'fd.Properties("Caption") = "On"
   td.Fields.Append fd
   
   '// Add The Classes
   '// ----------------
   For i = 1 To datTemp.intClassCount
      Set fd = td.CreateField(datTemp.strClasses(i), dbText)
         fd.Size = DAT_CLASS_TEXT_LENGTH
         td.Fields.Append fd
      
      '// Add Index
      Set ix = td.CreateIndex(datTemp.strClasses(i))
      Set fd = ix.CreateField(datTemp.strClasses(i), dbText)
         ix.Fields.Append fd
         td.Indexes.Append ix
   Next
   
   '// Add The Data Fields
   '// --------------------
   strSQL = "SELECT DISTINCTROW *" & CRLF()
   strSQL = strSQL & "FROM " & LPM_DAT_DEF_TABLE_NAME & "" & CRLF()
   strSQL = strSQL & "WHERE (((" & LPM_DAT_DEF_TABLE_NAME & ".DatType)='" & datTemp.strType & "'));"
   'Debug.Print strSQL
   
   Set rs = db.OpenRecordset(strSQL)
   Set rsClone = rs.Clone()
   rsClone.MoveFirst
   
   While rsClone.EOF = False
      If rsClone!FldActive Then
         Select Case rsClone!FldType
            Case "Number"
               lngFieldDataType = dbDouble
            Case "Currency"
               lngFieldDataType = dbCurrency
            Case "Text"
               lngFieldDataType = dbText
            Case "Boolean"
               lngFieldDataType = dbBoolean
         End Select
        
         Set fd = td.CreateField(rsClone!FldName, lngFieldDataType)
            If lngFieldDataType = dbText Then
               fd.Size = rsClone!FldSize
            End If
         td.Fields.Append fd
            
         If rsClone!FldIndex Then
            '// Add Index for Field
            Set ix = td.CreateIndex(rsClone!FldName)
               Set fd = ix.CreateField(rsClone!FldName, lngFieldDataType)
               ix.Fields.Append fd
            td.Indexes.Append ix
         End If
              
      End If
      rsClone.MoveNext
   Wend

   '// Append the Table Def to the DB
   db.TableDefs.Append td
   db.TableDefs.Refresh
   
   CreateDatTable = True

CreateDatTable_Done:
  Exit Function

CreateDatTable_Err:
Select Case Err
   Case 3010 '// Table Already Exists
      DeleteTable (DAT_TABLE_PRE & datTemp.strTable)
      Resume
   Case 3021 'No records
      CreateDatTable = False
      Resume CreateDatTable_Done
   Case 3211 '// Table Currently in Use
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & " Please close the table (" & DAT_TABLE_PRE & datTemp.strTable & ") and hit OK to continue"
      Resume
   Case 3219 '// Illegal Operation (Property Already Exists)
      Resume Next
   Case Else
      CreateDatTable = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CreateDatTable_Done
End Select

End Function



Function CreateRowTable(rowTemp As typRowType)
'//================================================================================//
'/|   FUNCTION: CreateRowTable                                                     |/
'/| PARAMETERS: rowTemp, the Row type to generate a table for.                     |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Creates a Table "zRowLogBal" whose name is based on the            |/
'/|             standard prefix constant and the entry in the "RowTypeTable" field |/
'/|             of table LPM_CONSTR_DEF_TABLE_NAME.                                |/
'/|    PURPOSE: Create table for storing individual LP vectors                     |/
'/|      USAGE: i = CreateRowTable(rowNew)                                         |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/10/97                                                            |/
'/|    HISTORY: 4/2/97   Changed to use the user defined type for row types        |/
'//================================================================================//

Dim db As Database
Dim td As TableDef
Dim fd As Field
Dim ix As Index
Dim i As Integer
Dim lngID As Long

On Error GoTo CreateRowTable_Err

   CreateRowTable = False
   
   Set db = CurrentDb                                                         '// Return Current Database.
   Set td = db.CreateTableDef(LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable)   '// Create New Table Definition
   
   '// RowType ID Field
   Set fd = td.CreateField(rowTemp.strType & "ID", dbLong)
      td.Fields.Append fd
   
   '// Add Primary Key Index
   Set ix = td.CreateIndex("PrimaryKey")
      ix.Primary = True
   Set fd = ix.CreateField(rowTemp.strType & "ID", dbLong)
      ix.Fields.Append fd
      td.Indexes.Append ix
   
   '//  RowType Code Field
   Set fd = td.CreateField(rowTemp.strType & "Code", dbText)
      fd.Size = 255
      td.Fields.Append fd
   
   '// Add The Classes
   '// ----------------
   For i = 1 To rowTemp.intClassCount
      '// Add Field
      Set fd = td.CreateField(rowTemp.strClasses(i), dbText)
         fd.Size = LPM_TEXT_LENGTH
      td.Fields.Append fd
      '// Add Index for Field
      Set ix = td.CreateIndex(rowTemp.strClasses(i))
      Set fd = ix.CreateField(rowTemp.strClasses(i), dbText)
         ix.Fields.Append fd
      td.Indexes.Append ix
   Next
   
   '// RHS Field
   Set fd = td.CreateField("RHS", dbDouble)
      td.Fields.Append fd
   
   '// RHS RANGE LO Field
   Set fd = td.CreateField("RHSLO", dbDouble)
      td.Fields.Append fd
      
   '// RHS RANGE UP Field
   Set fd = td.CreateField("RHSUP", dbDouble)
      td.Fields.Append fd
      
   '// RHS RANGE Field
   Set fd = td.CreateField("RANGE", dbDouble)
      td.Fields.Append fd
      
   '// SENSE Field
   Set fd = td.CreateField("SENSE", dbText)
      fd.Size = 1
      td.Fields.Append fd
   
   '// ACTIVITY Field
   Set fd = td.CreateField("ACTIVITY", dbDouble)
      td.Fields.Append fd
      
   '// SHADOW PRICE Field
   Set fd = td.CreateField("SHADOW", dbDouble)
      td.Fields.Append fd
      
   db.TableDefs.Append td
   db.TableDefs.Refresh
   
   CreateRowTable = True

CreateRowTable_Done:
  Exit Function

CreateRowTable_Err:
Select Case Err
   Case 3010 'Table Already Exists
      DeleteTable (LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable)
      Resume
   Case 3211 'Table Currently in Use
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & " Please close the table (" & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & ") and hit OK to continue"
      Resume
   Case 3219 'Illegal Operation (Property Already Exists)
      Resume Next
   Case Else
      CreateRowTable = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume CreateRowTable_Done
End Select

End Function

Function GenDatTbls()
'//================================================================================//
'/|   FUNCTION: GenDatTbls                                                         |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Calls the CreateDatTable() Routine                                 |/
'/|             for each Data Type in the Data Dictionary Tables                   |/
'/|    PURPOSE: Create tables for storing individual Data Records                  |/
'/|      USAGE: i= GenDatTbls()                                                    |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/10/97                                                            |/
'/|    HISTORY: 4/2/97                                                             |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset, rsClone As Recordset
Dim datType As typDatType
Dim i As Integer
Dim T As Date


On Error GoTo GenDatTbls_Err

GenDatTbls = False
T = Now()

Set db = CurrentDb

Set rs = db.OpenRecordset(LPM_DAT_DEF_TABLE_NAME)
Set rsClone = rs.Clone()
rsClone.MoveFirst
While rsClone.EOF = False
   datType = ReadDatType(rsClone!DatTypeID)
   If datType.intActive Then i = CreateDatTable(datType)
   rsClone.MoveNext
Wend

T = Now() - T
Debug.Print "GenDatTbls....." & Format(T, "hh:nn:ss")

GenDatTbls = True

   
GenDatTbls_Done:
  Exit Function

GenDatTbls_Err:
   GenDatTbls = False
   MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
   Resume GenDatTbls_Done
   Resume
End Function
Function GenElmTbls()
'//================================================================================//
'/|   FUNCTION: GenElmTbls                                                         |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Calls the CreateRowTable() and CreateColTable() Routines           |/
'/|             for each Row and Column Type in the Matrix Dictionary Tables       |/
'/|    PURPOSE: Create tables for storing individual LP vectors and rows           |/
'/|      USAGE: i= GenElmTbls()                                                    |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/10/97                                                            |/
'/|    HISTORY: 4/2/97                                                             |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset, rsClone As Recordset
Dim rowType As typRowType, colType As typColType
Dim i As Integer
Dim T As Date
Dim s As String

On Error GoTo GenElmTbls_Err

GenElmTbls = False
T = Now()

Set db = CurrentDb

'// R O W S //
Set rs = db.OpenRecordset(LPM_CONSTR_DEF_TABLE_NAME)
Set rsClone = rs.Clone()
rsClone.MoveFirst
While rsClone.EOF = False
   rowType = ReadRowType(rsClone!RowTypeID)
   If rowType.intActive Then
      i = CreateRowTable(rowType)
   End If
   rsClone.MoveNext
Wend

'// C O L U M N S //
Set rs = db.OpenRecordset(LPM_COLUMN_DEF_TABLE_NAME)
Set rsClone = rs.Clone()
rsClone.MoveFirst
While rsClone.EOF = False
   colType = ReadColType(rsClone!ColTypeID)
   If colType.intActive Then
      i = CreateColTable(colType)
   End If
   rsClone.MoveNext
Wend
   
T = Now() - T
s = "GenElmTbls....." & Format(T, "hh:nn:ss")
Debug.Print s
If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()


GenElmTbls = True

   
GenElmTbls_Done:
  Exit Function

GenElmTbls_Err:
   GenElmTbls = False
   MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
   Resume GenElmTbls_Done
   Resume
End Function

Function GenMtx()
'//================================================================================//
'/|   FUNCTION: GenMtx                                                             |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|             Calls the CreateRowTable() and CreateColTable() Routines           |/
'/|             for each Row and Column Type in the Matrix Dictionary Tables       |/
'/|    PURPOSE: Create tables for storing individual LP vectors and rows           |/
'/|      USAGE: i= GenMtx()                                                        |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/10/97                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database
Dim rs, rsClone As Recordset
Dim i As Integer
Dim strSQL As String
Dim strINS As String
Dim T As Date

On Error GoTo GenMtx_Err

GenMtx = False
T = Now()


Set db = CurrentDb

strINS = "INSERT INTO " & LPM_MATRIX_TABLE_NAME & CRLF() & _
         "  ( COL, ColID, ROW, RowID, COE )" & CRLF()

'// C O E F F I C I E N T S //
Set rs = db.OpenRecordset(LPM_COEFFS_DEF_TABLE_NAME)
Set rsClone = rs.Clone()
rsClone.MoveFirst
While rsClone.EOF = False
   strSQL = strINS & GenQry(rsClone!CoeffTypeID)
   DoCmd.SetWarnings False
   DoCmd.RunSQL strSQL
   Debug.Print rsClone!CoeffTypeID & "...." & Format(Now() - T, "hh:nn:ss") & " (Running)"
   DoCmd.SetWarnings True
   rsClone.MoveNext
Wend

   
T = Now() - T
Debug.Print "GenMtx...." & Format(T, "hh:nn:ss")
GenMtx = True
   
GenMtx_Done:
  Exit Function

GenMtx_Err:
   GenMtx = False
   MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
   Resume GenMtx_Done
   Resume
End Function


Function GenQry(lngCoeffID As Long) As String
'//================================================================================//
'/|   FUNCTION: GenQry                                                             |/
'/| PARAMETERS: lngCoeffID, ID from the Table LPM_COEFFS_DEF_TABLE_NAME            |/
'/|    RETURNS: SQL on Success and False by default or Failure                     |/
'/|    PURPOSE: Create SQL for getting a set of matrix coefficients for a given    |/
'/|             column type and row type.                                          |/
'/|      USAGE: i= GenQry(304)                                                     |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/15/97                                                            |/
'/|    HISTORY: 3/15/97  Originally Written                                        |/
'/|             3/17/97  Handle all 8 cases (Class in Col and/or Row and/or Qry)   |/
'/|             3/26/97  Add Column ID and Row ID                                  |/
'//================================================================================//

Dim db As Database, rs As Recordset, rsClone As Recordset, qdf As QueryDef, fld As Field

Dim i As Integer, j As Integer, intCols As Integer, intRows As Integer
Dim intIgNull As Integer, intFirstWhere As Integer, intColClassStatus As Integer
Dim intLast As Integer, intDupe As Integer, intCountWhere As Integer

Dim strColElTable As String, strRowElTable As String
Dim strColType As String, strRowType As String
Dim strQryName As String, strQryField As String, strSQL As String
Dim strColClassField As String, strRowClassField As String, strColClass As String, strRowClass As String
Dim Classes() As String

On Error GoTo GenQry_Err

GenQry = False
intIgNull = False
intFirstWhere = True
intCountWhere = 0
intRows = 0
intLast = 0
strColType = "XXX"
strRowType = "XXX"
strQryName = "XXX"
strQryField = "XXX"
strColElTable = "XXX"
strRowElTable = "XXX"
strColClassField = "XXX"
strColClass = "XXX"
   
   Set db = CurrentDb()
   
   '// G E T  C O L   A N D   R O W   I N F O R M A T I O N  //
   
   '// Get Col Type
   strColType = DLookup("ColType", LPM_COEFFS_DEF_TABLE_NAME, "CoeffTypeID = " & lngCoeffID)
   
   '// Get Row Type
   strRowType = DLookup("RowType", LPM_COEFFS_DEF_TABLE_NAME, "CoeffTypeID = " & lngCoeffID)
   
   '// Get Query Name
   strQryName = Brack(DLookup("CoeffRecSet", LPM_COEFFS_DEF_TABLE_NAME, "CoeffTypeID = " & lngCoeffID))
   
   '// Get Query's Field Name
   strQryField = Brack(DLookup("CoeffField", LPM_COEFFS_DEF_TABLE_NAME, "CoeffTypeID = " & lngCoeffID))
   
   '// Get Col Element Table Name
   strColElTable = Brack(LPM_COL_ELEMENT_TABLE_PRE & DLookup("ColTypeTable", LPM_COLUMN_DEF_TABLE_NAME, "ColType = " & "'" & strColType & "'"))
   
   '// Get Row Element Table Name
   strRowElTable = Brack(LPM_ROW_ELEMENT_TABLE_PRE & DLookup("RowTypeTable", LPM_CONSTR_DEF_TABLE_NAME, "RowType = " & "'" & strRowType & "'"))

   Set qdf = db.QueryDefs(UnBrack(strQryName))     '// Open the QueryDef Object
   intCols = CountColClasses(strColType)
   intRows = CountRowClasses(strRowType)
   
   ReDim Classes(intCols + intRows, 4)    '// 4 column array:  ClassName, Col?, Row?, Qry?
   
   For i = 1 To intCols + intRows
      Classes(i, 2) = "N"
      Classes(i, 3) = "N"
      Classes(i, 4) = "N"
   Next i
   
   '// Fill Array with the Column Classes
   For i = 1 To intCols
      strColClassField = "C" & CStr(i)
      strColClass = DLookup(strColClassField, LPM_COLUMN_DEF_TABLE_NAME, "ColType = " & "'" & strColType & "'")
      Classes(i, 1) = strColClass
      Classes(i, 2) = "Y"                 '// Put a 'Y' in the 2 column (Col?) of the array
      intLast = i
   Next i
      
   '// Search Rows for Classes
   For j = 1 To intRows
      strRowClassField = "R" & CStr(j)
      strRowClass = DLookup(strRowClassField, LPM_CONSTR_DEF_TABLE_NAME, "RowType = " & "'" & strRowType & "'")
      intDupe = False
         For i = 1 To intCols
            strColClassField = "C" & CStr(i)
            strColClass = DLookup(strColClassField, LPM_COLUMN_DEF_TABLE_NAME, "ColType = " & "'" & strColType & "'")
            If strRowClass = strColClass Then   '// If it is also a Column Class Then
               intDupe = True                   '// It's a duplicate
               Classes(i, 3) = "Y"              '// Put a 'Y' in the 3 column (Row?) of the array
            End If
         Next i
         
         If Not intDupe Then                          '// If this Row class is a new class Then
            Classes(intLast + 1, 1) = strRowClass     '// Add it to the array
            Classes(intLast + 1, 3) = "Y"             '// Put a 'Y' in the 3 column (Row?) of the array
            intLast = intLast + 1
         End If
   Next j
 
   '// Search The Query for Classes
   For i = 1 To intCols + intRows
      For j = 0 To qdf.Fields.Count - 1
         Set fld = qdf.Fields(j)
         If Classes(i, 1) = fld.Name Then
            Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
         End If
      Next j
   Next i
      
   '// Build the SQL statement (SELECT)
      
      strSQL = ""
      strSQL = "SELECT DISTINCTROW" & CRLF()       '// Select Distinct
      
      '// Field arguments to the Select Statement
      strSQL = strSQL & strColElTable & "." & strColType & "Code AS [COLUMN]," & CRLF()
      strSQL = strSQL & strColElTable & "." & strColType & "ID AS [CID]," & CRLF()        '// new
      strSQL = strSQL & strRowElTable & "." & strRowType & "Code AS [ROW]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & strRowType & "ID AS [RID]," & CRLF()        '// new
      strSQL = strSQL & strQryName & "." & strQryField & " AS [COEFF]" & CRLF()
      
      '// From Recordsets
      strSQL = strSQL & "FROM" & CRLF()
      strSQL = strSQL & strColElTable & "," & CRLF()
      strSQL = strSQL & strRowElTable & "," & CRLF()
      strSQL = strSQL & Brack(strQryName) & CRLF()

      intIgNull = True                             '// Set to ignore nulls in the following loops
   
      For i = 1 To intCols + intRows
         Select Case Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
            Case "YYY"     '// Class is in Column, Row, and Query
               intCountWhere = intCountWhere + 2
            Case "YYN"     '// Class is in Column and Row but not the Query
               intCountWhere = intCountWhere + 1
            Case "YNY"     '// Class is in Column and Query
               intCountWhere = intCountWhere + 1
            Case "NYY"     '// Class is in Row and Query
               intCountWhere = intCountWhere + 1
         End Select
      Next i
      
      
      If intCountWhere > 0 Then
      
         '// WHERE Statements
         strSQL = strSQL & "WHERE" & CRLF()
         strSQL = strSQL & "(" & CRLF()               '// Open All the Where's
   
         For i = 1 To intCols + intRows
            If Not intFirstWhere And Len(Classes(i, 1)) > 0 Then
               If Not ((Classes(i, 2) & Classes(i, 3) & Classes(i, 4) = "YNN") Or (Classes(i, 2) & Classes(i, 3) & Classes(i, 4) = "NYN")) Then
                  strSQL = strSQL & "AND" & CRLF()       '// Add 'AND' to the SQL if it's not the first Where Statement
               End If
            End If
            Select Case Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
               Case "YYY"     '// Class is in Column, Row, and Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strRowElTable & "." & Classes(i, 1) & ")" & CRLF()
                  strSQL = strSQL & "AND" & CRLF()
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "YYN"     '// Class is in Column and Row but not the Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strRowElTable & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "YNY"     '// Class is in Column and Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "NYY"     '// Class is in Row and Query
                  strSQL = strSQL & "(" & strRowElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case Else
            End Select
         Next i
         
         '// Close the WHERE Statement
         strSQL = strSQL & CRLF() & ")" & CRLF()
      
      End If

   
      '// Close the SELECT Statement
      strSQL = strSQL & ";"                           '// SQL Terminate Char

      '// DEBUG CODE
      'For i = 1 To intCols + intRows
      '   Debug.Print i & "   "; Classes(i, 1) & Space(15 - Len(Classes(i, 1))) & Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
      'Next i
      'Debug.Print ""
      'Debug.Print "SQL     = " & CRLF() & strSQL
      'Debug.Print ""
      'Debug.Print "ColType = " & strColType
      'Debug.Print "RowType = " & strRowType
      'Debug.Print "Query   = " & strQryName
      'Debug.Print "Field   = " & strQryField
      'Debug.Print "Col Tbl = " & strColElTable
      'Debug.Print "Row Tbl = " & strRowElTable
      'Debug.Print "Wheres  = " & intCountWhere
   
      GenQry = strSQL
   
GenQry_Done:
  Exit Function

GenQry_Err:
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         GenQry = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume GenQry_Done
      End If
   Case 3265
      GenQry = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query and field name exist."
      Resume GenQry_Done
    Case Else
      GenQry = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume GenQry_Done
      Resume
End Select
   
End Function


Function InitModel()
'//================================================================================//
'/|   FUNCTION: InitModel                                                          |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Clean out all Element tables                                       |/
'/|      USAGE: i= InitModel()                                                     |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/26/97                                                            |/
'/|    HISTORY: 3/26/97  Originally Written                                        |/
'//================================================================================//
   
   Dim DefaultWorkspace As Workspace
   Dim MyDatabase As Database, MyTableDef As TableDef, MyQueryDef As QueryDef, MyField As Field
   Dim strSQL As String
   Dim i As Integer, j As Integer

   On Error GoTo InitModel_Err

   InitModel = False
   DoCmd.SetWarnings False

   Set MyDatabase = CurrentDb
    
   '// TABLE LOOP //
   For i = 0 To MyDatabase.TableDefs.Count - 1
      
      Set MyTableDef = MyDatabase.TableDefs(i)
      
      '// Delete all Element Tables //
      If InStr(MyTableDef.Name, LPM_ROW_ELEMENT_TABLE_PRE) Or InStr(MyTableDef.Name, LPM_COL_ELEMENT_TABLE_PRE) > 0 Then
         DeleteTable (MyTableDef.Name)
      End If
   
   Next i

   i = InitMtx()

   DoCmd.SetWarnings True
   InitModel = True
   
InitModel_Done:
   Exit Function

InitModel_Err:
   InitModel = False
   MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
   Resume InitModel_Done
   Resume

End Function

Function InitMtx()

   Dim strSQL As String

   On Error GoTo InitMtx_Err

   InitMtx = False
   DoCmd.SetWarnings False

   '// Delete all Records from LPM_MATRIX_... Tables //
   strSQL = "Delete * FROM " & LPM_MATRIX_TABLE_NAME
   DoCmd.RunSQL strSQL
   strSQL = "Delete * FROM " & LPM_COLUMN_TABLE_NAME
   DoCmd.RunSQL strSQL
   strSQL = "Delete * FROM " & LPM_CONSTR_TABLE_NAME
   DoCmd.RunSQL strSQL
   strSQL = "Delete * FROM " & LPM_CONSTR_TABLE_NAME_IMPORT
   DoCmd.RunSQL strSQL
   strSQL = "Delete * FROM " & LPM_COLUMN_TABLE_NAME_IMPORT
   DoCmd.RunSQL strSQL
   
   
   DoCmd.SetWarnings True
   InitMtx = True
     
InitMtx_Done:
   Exit Function

InitMtx_Err:
   InitMtx = False
   MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
   Resume InitMtx_Done
   Resume

End Function


Function MakeMatGenQry(ID As Long) As String
'//================================================================================//
'/|   FUNCTION: MakeMatGenQry                                                      |/
'/| PARAMETERS: lngCoeffID, ID from the Table LPM_COEFFS_DEF_TABLE_NAME            |/
'/|    RETURNS: SQL on Success and False by default or Failure                     |/
'/|    PURPOSE: Create SQL for getting a set of matrix coefficients for a given    |/
'/|             column type and row type.                                          |/
'/|      USAGE: i= MakeMatGenQry(23)                                               |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/15/97                                                            |/
'/|    HISTORY: 3/15/97  Originally Written                                        |/
'/|             3/17/97  Handle all 8 cases (Class in Col and/or Row and/or Qry)   |/
'/|             3/26/97  Add Column ID and Row ID                                  |/
'/|             4/11/97  Change Modularity based on Read, Make Qry, Execute        |/
'/|                      also allow coeff recset to access tables as well as qry's |/
'//================================================================================//

Dim rowTemp As typRowType
Dim colTemp As typColType
Dim coeffTemp As typCoeffType
Dim db As Database
Dim qdf As QueryDef, fld As Field, tdf As TableDef

Dim i As Integer, j As Integer, intCols As Integer, intRows As Integer
Dim intIgNull As Integer, intFirstWhere As Integer, intColClassStatus As Integer
Dim intLast As Integer, intDupe As Integer, intCountWhere As Integer
Dim intFirst3265Error As Integer
Dim strQorT As String
Dim strColElTable As String, strRowElTable As String
Dim strColType As String, strRowType As String
Dim strQryName As String, strQryField As String, strSQL As String
Dim strColClass As String, strRowClass As String
Dim Classes() As String

On Error GoTo MakeMatGenQry_Err

MakeMatGenQry = False
intIgNull = False
intFirstWhere = True
intCountWhere = 0
intRows = 0
intLast = 0
strQorT = "Q"
intFirst3265Error = True

coeffTemp = ReadCoeffType(ID)
colTemp = ReadColType(coeffTemp.lngColID)
rowTemp = ReadRowType(coeffTemp.lngRowID)


Set db = CurrentDb


   
   '// G E T  C O L   A N D   R O W   I N F O R M A T I O N  //
   If colTemp.intActive * rowTemp.intActive * coeffTemp.intActive = 0 Then
      gMatGenOK = False
      MakeMatGenQry = False
      GoTo MakeMatGenQry_Done
   End If
      
   '// Get Col Type
   strColType = coeffTemp.strColType
      
   '// Get Row Type
   strRowType = coeffTemp.strRowType
   
   
   strQryName = Brack(coeffTemp.strRecSet)                              '// Get Query Name
   strQryField = Brack(coeffTemp.strCoeffFld)                           '// Get Query's Field Name
   strColElTable = Brack(LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable)  '// Get Col Element Table Name
   strRowElTable = Brack(LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable)  '// Get Row Element Table Name
   Set qdf = db.QueryDefs(UnBrack(strQryName))                          '// Open the QueryDef Object
   intCols = colTemp.intClassCount
   intRows = rowTemp.intClassCount
   
   ReDim Classes(intCols + intRows, 4)    '// 4 column array:  ClassName, Col?, Row?, Qry?
   
   For i = 1 To intCols + intRows
      Classes(i, 2) = "N"
      Classes(i, 3) = "N"
      Classes(i, 4) = "N"
   Next i
   
   '// Fill Array with the Column Classes
   For i = 1 To intCols
      strColClass = colTemp.strClasses(i)
      Classes(i, 1) = strColClass
      Classes(i, 2) = "Y"                 '// Put a 'Y' in the 2 column (Col?) of the array
      intLast = i
   Next i

   '// Search Rows for Classes
   For j = 1 To intRows
      strRowClass = rowTemp.strClasses(j)
      intDupe = False
         For i = 1 To intCols
            strColClass = colTemp.strClasses(i)
            If strRowClass = strColClass Then   '// If it is also a Column Class Then
               intDupe = True                   '// It's a duplicate
               Classes(i, 3) = "Y"              '// Put a 'Y' in the 3 column (Row?) of the array
            End If
         Next i
         
         If Not intDupe Then                          '// If this Row class is a new class Then
            Classes(intLast + 1, 1) = strRowClass     '// Add it to the array
            Classes(intLast + 1, 3) = "Y"             '// Put a 'Y' in the 3 column (Row?) of the array
            intLast = intLast + 1
         End If
   Next j
 
   
   Select Case strQorT
   
   Case "Q"
      '// Search The Query for Classes
      For i = 1 To intCols + intRows
         For j = 0 To qdf.Fields.Count - 1
            Set fld = qdf.Fields(j)
            If Classes(i, 1) = fld.Name Then
               Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
            End If
         Next j
      Next i
      
   Case "T"
      '// Search The Table for Classes
      For i = 1 To intCols + intRows
         For j = 0 To tdf.Fields.Count - 1
            Set fld = tdf.Fields(j)
            If Classes(i, 1) = fld.Name Then
               Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
            End If
         Next j
      Next i
      
   End Select
      
   '// Build the SQL statement (SELECT)
      
      strSQL = ""
      strSQL = "SELECT DISTINCTROW" & CRLF()       '// Select Distinct
      
      '// Field arguments to the Select Statement
      strSQL = strSQL & strColElTable & "." & strColType & "Code AS [COLUMN]," & CRLF()
      strSQL = strSQL & strColElTable & "." & strColType & "ID AS [CID]," & CRLF()        '// new
      strSQL = strSQL & strColElTable & "." & "OBJ AS [OBJ]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "LO AS [LO]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "UP AS [UP]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "FREE AS [FREE]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "SOS AS [SOS]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & strRowType & "Code AS [ROW]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & strRowType & "ID AS [RID]," & CRLF()        '// new
      strSQL = strSQL & strRowElTable & "." & "SENSE AS [SENSE]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & "RHS AS [RHS]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & "RANGE AS [RANGE]," & CRLF()
      strSQL = strSQL & strQryName & "." & strQryField & " AS [COEFF]" & CRLF()
      
      '// From Recordsets
      strSQL = strSQL & "FROM" & CRLF()
      strSQL = strSQL & strColElTable & "," & CRLF()
      strSQL = strSQL & strRowElTable & "," & CRLF()
      strSQL = strSQL & Brack(strQryName) & CRLF()

      intIgNull = True                             '// Set to ignore nulls in the following loops
   
      For i = 1 To intCols + intRows
         Select Case Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
            Case "YYY"     '// Class is in Column, Row, and Query
               intCountWhere = intCountWhere + 2
            Case "YYN"     '// Class is in Column and Row but not the Query
               intCountWhere = intCountWhere + 1
            Case "YNY"     '// Class is in Column and Query
               intCountWhere = intCountWhere + 1
            Case "NYY"     '// Class is in Row and Query
               intCountWhere = intCountWhere + 1
         End Select
      Next i
      
      
      If intCountWhere > 0 Then
      
         '// WHERE Statements
         strSQL = strSQL & "WHERE" & CRLF()
         strSQL = strSQL & "(" & CRLF()               '// Open All the Where's
   
         For i = 1 To intCols + intRows
            If Not intFirstWhere And Len(Classes(i, 1)) > 0 Then
               If Not ((Classes(i, 2) & Classes(i, 3) & Classes(i, 4) = "YNN") Or (Classes(i, 2) & Classes(i, 3) & Classes(i, 4) = "NYN")) Then
                  strSQL = strSQL & "AND" & CRLF()       '// Add 'AND' to the SQL if it's not the first Where Statement
               End If
            End If
            Select Case Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
               Case "YYY"     '// Class is in Column, Row, and Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strRowElTable & "." & Classes(i, 1) & ")" & CRLF()
                  strSQL = strSQL & "AND" & CRLF()
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "YYN"     '// Class is in Column and Row but not the Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strRowElTable & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "YNY"     '// Class is in Column and Query
                  strSQL = strSQL & "(" & strColElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case "NYY"     '// Class is in Row and Query
                  strSQL = strSQL & "(" & strRowElTable & "." & Classes(i, 1) & " = " & strQryName & "." & Classes(i, 1) & ")" & CRLF()
                  intFirstWhere = False
               Case Else
            End Select
         Next i
         
         '// Close the WHERE Statement
         strSQL = strSQL & CRLF() & ")" & CRLF()
      
      End If

   
      '// Close the SELECT Statement
      strSQL = strSQL & ";"                           '// SQL Terminate Char

      '// DEBUG CODE
      'For i = 1 To intCols + intRows
      '   Debug.Print i & "   "; Classes(i, 1) & Space(15 - Len(Classes(i, 1))) & Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
      'Next i
      'Debug.Print ""
      'Debug.Print "SQL     = " & CRLF() & strSQL
      'Debug.Print ""
      'Debug.Print "ColType = " & strColType
      'Debug.Print "RowType = " & strRowType
      'Debug.Print "Query   = " & strQryName
      'Debug.Print "Field   = " & strQryField
      'Debug.Print "Col Tbl = " & strColElTable
      'Debug.Print "Row Tbl = " & strRowElTable
      'Debug.Print "Wheres  = " & intCountWhere
   
      MakeMatGenQry = strSQL
   
MakeMatGenQry_Done:
  Exit Function

MakeMatGenQry_Err:
If Len(strQryName) < 1 Then
   Debug.Print "           No CoeffRecSet - Block Not Submitted to matrix"
   gMatGenOK = False
   MakeMatGenQry = False
   Resume MakeMatGenQry_Done
   Resume
End If
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         MakeMatGenQry = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume MakeMatGenQry_Done
      End If
   Case 3265
      If intFirst3265Error Then
         Set tdf = db.TableDefs(UnBrack(strQryName))                          '// Open the TableDef Object
         intFirst3265Error = False
         strQorT = "T"
         Resume Next
      Else
         MakeMatGenQry = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query (or table) and field name exist."
         Resume MakeMatGenQry_Done
      End If
    Case Else
      MakeMatGenQry = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume MakeMatGenQry_Done
      Resume
End Select
   
End Function

Function MakeMatGenQry2(ID As Long) As String
'//================================================================================//
'/|   FUNCTION: MakeMatGenQry2                                                     |/
'/| PARAMETERS: lngCoeffID, ID from the Table LPM_COEFFS_DEF_TABLE_NAME            |/
'/|    RETURNS: SQL on Success and False by default or Failure                     |/
'/|    PURPOSE: Create SQL for getting a set of matrix coefficients for a given    |/
'/|             column type and row type.                                          |/
'/|      USAGE: i= MakeMatGenQry2(23)                                              |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/15/97                                                            |/
'/|    HISTORY: 3/15/97  Originally Written                                        |/
'/|             3/17/97  Handle all 8 cases (Class in Col and/or Row and/or Qry)   |/
'/|             3/26/97  Add Column ID and Row ID                                  |/
'/|             4/11/97  Change Modularity based on Read, Make Qry, Execute        |/
'/|                      also allow coeff recset to access tables as well as qry's |/
'/|             7/22/99  Change the join implementation from WHERE clauses to      |/
'/|                      INNER JOIN syntax                                         |/
'/|             11/8/99  Complete and test the previous change.  Get all 27 unique |/
'/|                      JOIN Combo Possibilites, translate to  the 11 general     |/
'/|                      query structures                                          |/
'/|             8/2/00                                                             |/
'//================================================================================//

Dim rowTemp As typRowType
Dim colTemp As typColType
Dim coeffTemp As typCoeffType
Dim db As Database
Dim qdf As QueryDef, fld As Field, tdf As TableDef

Dim i As Integer, j As Integer, intCols As Integer, intRows As Integer
Dim intIgNull As Integer, intFirstWhere As Integer, intColClassStatus As Integer
Dim intLast As Integer, intDupe As Integer, intCountWhere As Integer
Dim intCountC2R As Integer, intCountC2Q As Integer, intCountR2Q As Integer
Dim intFirst3265Error As Integer
Dim strQorT As String
Dim strColElTable As String, strRowElTable As String
Dim strColType As String, strRowType As String
Dim strQryName As String, strQryField As String, strSQL As String
Dim strColClass As String, strRowClass As String
Dim strClassIncCode As String
Dim strClassName As String
Dim strC2Rcode As String, strC2Qcode As String, strR2Qcode As String
Dim strJoinComboType As String
Dim intMatgenQryType As Integer
Dim sF00 As String, sF01 As String, sF02 As String, sF03 As String, sF04 As String, sF05 As String
Dim sJ1 As String, sJ2 As String, sJ3 As String
Dim sC01 As String, sC02 As String, sC03 As String

Dim Classes() As String

On Error GoTo MakeMatGenQry2_Err

MakeMatGenQry2 = False
intIgNull = False
intFirstWhere = True
intCountWhere = 0
intCountC2R = 0
intCountC2Q = 0
intCountR2Q = 0
intRows = 0
intLast = 0
strQorT = "Q"
strJoinComboType = ""
intFirst3265Error = True

coeffTemp = ReadCoeffType(ID)
colTemp = ReadColType(coeffTemp.lngColID)
rowTemp = ReadRowType(coeffTemp.lngRowID)


Set db = CurrentDb


   
   '// G E T  C O L   A N D   R O W   I N F O R M A T I O N  //
   If colTemp.intActive * rowTemp.intActive * coeffTemp.intActive = 0 Then
      gMatGenOK = False
      MakeMatGenQry2 = False
      GoTo MakeMatGenQry2_Done
   End If
      
   '// Get Col Type
   strColType = coeffTemp.strColType
      
   '// Get Row Type
   strRowType = coeffTemp.strRowType
   
   
   strQryName = Brack(coeffTemp.strRecSet)                              '// Get Query Name
   strQryField = Brack(coeffTemp.strCoeffFld)                           '// Get Query's Field Name
   strColElTable = Brack(LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable)  '// Get Col Element Table Name
   strRowElTable = Brack(LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable)  '// Get Row Element Table Name
   Set qdf = db.QueryDefs(UnBrack(strQryName))                          '// Open the QueryDef Object
   intCols = colTemp.intClassCount
   intRows = rowTemp.intClassCount
   
   ReDim Classes(intCols + intRows, 4)    '// 4 column array:  ClassName, Col?, Row?, Qry?
   
   '// Initialize the Class Array with all 'N'
   For i = 1 To intCols + intRows
      Classes(i, 2) = "N"
      Classes(i, 3) = "N"
      Classes(i, 4) = "N"
   Next i
   
   '// Fill Array with the Column Classes
   For i = 1 To intCols
      strColClass = colTemp.strClasses(i)
      Classes(i, 1) = strColClass
      Classes(i, 2) = "Y"                 '// Put a 'Y' in the 2 column (Col?) of the array
      intLast = i
   Next i

   '// Search Rows for Classes
   For j = 1 To intRows
      strRowClass = rowTemp.strClasses(j)
      intDupe = False
         For i = 1 To intCols
            strColClass = colTemp.strClasses(i)
            If strRowClass = strColClass Then   '// If it is also a Column Class Then
               intDupe = True                   '// It's a duplicate
               Classes(i, 3) = "Y"              '// Put a 'Y' in the 3 column (Row?) of the array
            End If
         Next i
         
         If Not intDupe Then                          '// If this Row class is a new class Then
            Classes(intLast + 1, 1) = strRowClass     '// Add it to the array
            Classes(intLast + 1, 3) = "Y"             '// Put a 'Y' in the 3 column (Row?) of the array
            intLast = intLast + 1
         End If
   Next j
 
   
   Select Case strQorT
   
   Case "Q"
      '// Search The Query for Classes
      For i = 1 To intCols + intRows
         For j = 0 To qdf.Fields.Count - 1
            Set fld = qdf.Fields(j)
            If Classes(i, 1) = fld.Name Then
               Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
            End If
         Next j
      Next i
      
   Case "T"
      '// Search The Table for Classes
      For i = 1 To intCols + intRows
         For j = 0 To tdf.Fields.Count - 1
            Set fld = tdf.Fields(j)
            If Classes(i, 1) = fld.Name Then
               Classes(i, 4) = "Y"                       '// Put a 'Y' in the 4 column (Qry?) of the array
            End If
         Next j
      Next i
      
   End Select
      
   '// Build the SQL statement (SELECT)
      
      strSQL = ""
      strSQL = "SELECT DISTINCT" & CRLF()       '// Select Distinct {Changed from DISTINCTROW 11/7/03  --STM}
      
      '// Field arguments to the Select Statement
      strSQL = strSQL & strColElTable & "." & strColType & "Code AS [COLUMN]," & CRLF()
      strSQL = strSQL & strColElTable & "." & strColType & "ID AS [CID]," & CRLF()        '// new
      strSQL = strSQL & strColElTable & "." & "OBJ AS [OBJ]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "LO AS [LO]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "UP AS [UP]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "FREE AS [FREE]," & CRLF()
      strSQL = strSQL & strColElTable & "." & "SOS AS [SOS]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & strRowType & "Code AS [ROW]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & strRowType & "ID AS [RID]," & CRLF()        '// new
      strSQL = strSQL & strRowElTable & "." & "SENSE AS [SENSE]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & "RHS AS [RHS]," & CRLF()
      strSQL = strSQL & strRowElTable & "." & "RANGE AS [RANGE]," & CRLF()
      strSQL = strSQL & strQryName & "." & strQryField & " AS [COEFF]" & CRLF()
      
      
      '*&*&*&*&*&*&*&*&*
      
      ''// From Recordsets
      'strSQL = strSQL & "FROM" & CRLF()
      'strSQL = strSQL & strColElTable & "," & CRLF()
      'strSQL = strSQL & strRowElTable & "," & CRLF()
      'strSQL = strSQL & Brack(strQryName) & CRLF()

      intIgNull = True                             '// Set to ignore nulls in the following loops
   
      For i = 1 To intCols + intRows
         strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
         Select Case strClassIncCode
            Case "YYY"     '// Class is in Column, Row, and Query
               intCountWhere = intCountWhere + 2
               intCountC2R = intCountC2R + 1
               intCountC2Q = intCountC2Q + 1
            Case "YYN"     '// Class is in Column and Row but not the Query
               intCountWhere = intCountWhere + 1
               intCountC2R = intCountC2R + 1
            Case "YNY"     '// Class is in Column and Query
               intCountWhere = intCountWhere + 1
               intCountC2Q = intCountC2Q + 1
            Case "NYY"     '// Class is in Row and Query
               intCountWhere = intCountWhere + 1
               intCountR2Q = intCountR2Q + 1
         End Select
      Next i
      
      
      '*((*(*(*(*(*(*(*(
      '// Change all the counts that are greater than one to "2"
      '// and convert the rest to strings
      If intCountC2R >= 2 Then
         strC2Rcode = "2"
      Else
         strC2Rcode = CStr(intCountC2R)
      End If
               
      If intCountC2Q >= 2 Then
         strC2Qcode = "2"
      Else
         strC2Qcode = CStr(intCountC2Q)
      End If
               
      If intCountR2Q >= 2 Then
         strR2Qcode = "2"
      Else
         strR2Qcode = CStr(intCountR2Q)
      End If
      
      strJoinComboType = strC2Rcode & strC2Qcode & strR2Qcode
      
      
      '// convert the join combo type (27 unique possibilities) to one of the 11 (0-10) types of generic block queries
      Select Case strJoinComboType
         Case "000"
            intMatgenQryType = 0
         Case "010", "020"
            intMatgenQryType = 1
         Case "001", "002"
            intMatgenQryType = 2
         Case "100", "200"
            intMatgenQryType = 3
         Case "110", "120"
            intMatgenQryType = 4
         Case "101", "102"
            intMatgenQryType = 5
         Case "111", "122", "121", "112", "211", "222", "221", "212"
            intMatgenQryType = 6
         Case "210", "220"
            intMatgenQryType = 7
         Case "201"
            intMatgenQryType = 8
         Case "011", "022", "021", "012"
            intMatgenQryType = 9
         Case "202"
            intMatgenQryType = 10
      End Select
      
      
      
      
      '// The Beginning of the FROM clause
      '===================================================================
      sF00 = "FROM " & strColElTable & ", " & strRowElTable & ", " & strQryName & CRLF()
      sF01 = "FROM " & strRowElTable & ", " & strColElTable & CRLF()
      sF02 = "FROM " & strColElTable & ", " & strRowElTable & CRLF()
      sF03 = "FROM " & strColElTable & CRLF()
      sF04 = "FROM " & strQryName & ", " & strColElTable & CRLF()
      sF05 = "FROM (" & strColElTable & CRLF()
      
      '// The INNER JOIN clauses
      '===================================================================
      sJ1 = "INNER JOIN " & strQryName & " ON" & CRLF()
      sJ2 = "INNER JOIN (" & strRowElTable & CRLF()
      sJ3 = "INNER JOIN " & strRowElTable & " ON" & CRLF()
      
      sC01 = ""
      sC02 = ""
      sC03 = ""
      
      '// The Field Joins
      '===================================================================
      
         '// Column to Row (aka C2R and C01)
         '===================================
         intFirstWhere = True
         If intCountC2R > 0 Then
         
            '// "WHERE" Statements
            'sC01 = sC01 & "(" & CRLF()
               
            For i = 1 To intCols + intRows
               strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
               strClassName = Classes(i, 1)
               If Not intFirstWhere And Len(Classes(i, 1)) > 0 Then
                  If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                     If Not ((strClassIncCode = "YNY") Or (strClassIncCode = "NYY")) Then '//Class is not in the C2R join
                        sC01 = sC01 & "AND" & CRLF()       '// Add 'AND' to the SQL if it's not the first Where Statement
                     End If
                  End If
               End If
               Select Case strClassIncCode
                  Case "YYY"     '// Class is in Column, Row, and Query
                     sC01 = sC01 & "(" & strColElTable & "." & strClassName & " = " & strRowElTable & "." & strClassName & ")" & CRLF()
                     intFirstWhere = False
                  Case "YYN"     '// Class is in Column and Row but not the Query
                     sC01 = sC01 & "(" & strColElTable & "." & strClassName & " = " & strRowElTable & "." & strClassName & ")" & CRLF()
                     intFirstWhere = False
                  Case Else
               End Select
            Next i
            
            '// Close the WHERE Statement
            'sC01 = sC01 & CRLF() & ")" & CRLF()
         
         End If
      
      
         '// Column to Qry (aka C2Q and C02)
         '===================================
         intFirstWhere = True
         If intCountC2Q > 0 Then
         
            '// "WHERE" Statements
            'sC02 = sC02 & "(" & CRLF()
               
            For i = 1 To intCols + intRows
               strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
               strClassName = Classes(i, 1)
               If Not intFirstWhere And Len(Classes(i, 1)) > 0 Then
                  If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                     If Not ((strClassIncCode = "YYN") Or (strClassIncCode = "NYY")) Then '//Class is not in the C2Q join
                        sC02 = sC02 & "AND" & CRLF()       '// Add 'AND' to the SQL if it's not the first Where Statement
                     End If
                  End If
               End If
               Select Case strClassIncCode
                  Case "YYY"     '// Class is in Column, Row, and Query
                     sC02 = sC02 & "(" & strColElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & CRLF()
                     intFirstWhere = False
                  Case "YNY"     '// Class is in Column and Query
                     sC02 = sC02 & "(" & strColElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & CRLF()
                     intFirstWhere = False
                  Case Else
               End Select
            Next i
            
            '// Close the WHERE Statement
            'sC02 = sC02 & CRLF() & ")" & CRLF()
         
         End If
      
      
         '// Row to Qry (aka R2Q and C03)
         '===================================
         intFirstWhere = True
         If intCountR2Q > 0 Then
         
            '// "WHERE" Statements
            'sC03 = sC03 & "(" & CRLF()
               
            For i = 1 To intCols + intRows
               strClassIncCode = Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
               strClassName = Classes(i, 1)
               If Not intFirstWhere And Len(Classes(i, 1)) > 0 Then
                  If Not ((strClassIncCode = "YNN") Or (strClassIncCode = "NYN")) Then    '//Class is in more than one place
                     If Not ((strClassIncCode = "YYY") Or (strClassIncCode = "YYN") Or (strClassIncCode = "YNY")) Then '//Class is not in the R2Q join
                        sC03 = sC03 & "AND" & CRLF()       '// Add 'AND' to the SQL if it's not the first Where Statement
                     End If
                  End If
               End If
               Select Case strClassIncCode
                  Case "NYY"     '// Class is in Row and Query
                     sC03 = sC03 & "(" & strRowElTable & "." & strClassName & " = " & strQryName & "." & strClassName & ")" & CRLF()
                     intFirstWhere = False
                  Case Else
               End Select
            Next i
            
            '// Close the WHERE Statement
            'sC03 = sC03 & CRLF() & ")" & CRLF()
         
         End If
      
      Select Case intMatgenQryType     '// note that case 7 and case 4 are the same
                                       '//                8 and      5 too
         Case 0
            strSQL = strSQL & sF00 & ";"
         Case 1
            strSQL = strSQL & sF01 & sJ1 & sC02 & ";"
         Case 2
            strSQL = strSQL & sF02 & sJ1 & sC03 & ";"
         Case 3
            strSQL = strSQL & sF04 & sJ3 & sC01 & ";"
         Case 4
            strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & CRLF() & sJ1 & sC02 & ";"
         Case 5
            strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & CRLF() & sJ1 & sC03 & ";"
         Case 6
            strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & CRLF() & sJ1 & sC02 & " AND " & CRLF() & sC03 & ";"
         Case 7
            strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & CRLF() & sJ1 & sC02 & ";"
         Case 8
            strSQL = strSQL & sF05 & sJ3 & sC01 & ") " & CRLF() & sJ1 & sC03 & ";"
         Case 9
            strSQL = strSQL & sF03 & sJ2 & CRLF() & sJ1 & sC03 & ") ON" & CRLF() & sC02 & ";"
         Case 10
            strSQL = strSQL & sF03 & sJ2 & CRLF() & sJ1 & sC03 & ") ON" & CRLF() & sC01 & ";"
      End Select
      
      
      
      
      '// DEBUG CODE
      'For i = 1 To intCols + intRows
      '   Debug.Print i & "   "; Classes(i, 1) & Space(15 - Len(Classes(i, 1))) & Classes(i, 2) & Classes(i, 3) & Classes(i, 4)
     ' Next i
     ' Debug.Print ""
     ' Debug.Print "SQL     = " & CRLF() & strSQL
     ' Debug.Print ""
     ' Debug.Print "ColType = " & strColType
     ' Debug.Print "RowType = " & strRowType
     ' Debug.Print "Query   = " & strQryName
     ' Debug.Print "Field   = " & strQryField
     ' Debug.Print "Col Tbl = " & strColElTable
     ' Debug.Print "Row Tbl = " & strRowElTable
     ' Debug.Print "Wheres  = " & intCountWhere
     ' Debug.Print "JCombo  = " & strJoinComboType
     ' Debug.Print "QryType = " & intMatgenQryType
   
      MakeMatGenQry2 = strSQL
   
MakeMatGenQry2_Done:
  Exit Function

MakeMatGenQry2_Err:
If Len(strQryName) < 1 Then
   Debug.Print "           No CoeffRecSet - Block Not Submitted to matrix"
   gMatGenOK = False
   MakeMatGenQry2 = False
   Resume MakeMatGenQry2_Done
   Resume
End If
Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         MakeMatGenQry2 = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume MakeMatGenQry2_Done
      End If
   Case 3265
      If intFirst3265Error Then
         Set tdf = db.TableDefs(UnBrack(strQryName))                          '// Open the TableDef Object
         intFirst3265Error = False
         strQorT = "T"
         Resume Next
      Else
         MakeMatGenQry2 = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query (or table) and field name exist."
         Resume MakeMatGenQry2_Done
      End If
    Case Else
      MakeMatGenQry2 = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume MakeMatGenQry2_Done
      Resume
End Select
End Function


Function MakePopColQry(colTemp As typColType) As String

Dim strINS As String, strSQL As String, strSelObj As String, strVectorName As String
Dim strColClass As String
Dim j As Long

   MakePopColQry = "FALSE"

   strINS = "INSERT INTO " & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & " ( " & colTemp.strType & "ID, "
   strSQL = strINS
   strSelObj = ""
   strVectorName = "'" & colTemp.strPrefix & "' "
   
   For j = 1 To colTemp.intClassCount
      strSQL = strSQL & colTemp.strClasses(j) & ", "
      strColClass = colTemp.strClasses(j)
      strSelObj = strSelObj & colTemp.strRecSet & "." & strColClass & " AS " & strColClass & ", " & CRLF()
      strVectorName = strVectorName & Chr(38) & " " & "'" & "_" & "'" & " " & Chr(38) & " " & Brack(strColClass)
   Next j
   
   If Len(colTemp.strOBJFld) > 0 Then
      strSQL = strSQL & "OBJ, "
   End If
   
   If Len(colTemp.strLOFld) > 0 Then
      strSQL = strSQL & "LO, "
   End If
   
   If Len(colTemp.strUPFld) > 0 Then
      strSQL = strSQL & "UP, "
   End If
   
   If Len(colTemp.strSOS) > 0 Then
      strSQL = strSQL & "SOS, "
   End If
   
   If colTemp.intFREE < 0 Then
      strSQL = strSQL & "FREE, "
   End If
   
   strSQL = strSQL & colTemp.strType & "Code )" & CRLF()
   strSQL = strSQL & "SELECT DISTINCTROW" & CRLF() & "QCntr(" & strVectorName & ") AS ID," & CRLF()
   strSQL = strSQL & strSelObj
   
   If Len(colTemp.strOBJFld) > 0 Then
      strSQL = strSQL & colTemp.strRecSet & "." & colTemp.strOBJFld & " AS OBJ, " & CRLF()
   End If
   
   If Len(colTemp.strLOFld) > 0 Then
      strSQL = strSQL & colTemp.strRecSet & "." & colTemp.strLOFld & " AS LO, " & CRLF()
   End If
   
   If Len(colTemp.strUPFld) > 0 Then
      strSQL = strSQL & colTemp.strRecSet & "." & colTemp.strUPFld & " AS UP, " & CRLF()
   End If
   
   If Len(colTemp.strSOS) > 0 Then
      strSQL = strSQL & "'" & colTemp.strSOS & "'" & " AS SOS, " & CRLF()
   End If
   
   If colTemp.intFREE < 0 Then
      strSQL = strSQL & colTemp.intFREE & " AS FREE, " & CRLF()
   End If

   strVectorName = strVectorName & " AS VectorName" & CRLF()
   
   strSQL = strSQL & strVectorName
   
   strSQL = strSQL & "FROM " & colTemp.strRecSet & CRLF()
   strSQL = strSQL & ";"
   
   MakePopColQry = strSQL

End Function

Function MakePopRowQry(rowTemp As typRowType) As String

Dim strINS As String, strSQL As String, strSelObj As String, strConstraintName As String
Dim strRowClass As String
Dim j As Long

   MakePopRowQry = "FALSE"

   strINS = "INSERT INTO " & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & " ( " & rowTemp.strType & "ID, "
   strSQL = strINS
   strSelObj = ""
   strConstraintName = "'" & rowTemp.strPrefix & "' "
   
   For j = 1 To rowTemp.intClassCount
      strSQL = strSQL & rowTemp.strClasses(j) & ", "
      strRowClass = rowTemp.strClasses(j)
      strSelObj = strSelObj & rowTemp.strRecSet & "." & strRowClass & " AS " & strRowClass & ", " & CRLF()
      strConstraintName = strConstraintName & Chr(38) & " " & "'" & "_" & "'" & " " & Chr(38) & " " & Brack(strRowClass)
   Next j
   
   If Len(rowTemp.strRHSFld) > 0 Then
      strSQL = strSQL & "RHS, "
   End If
   If Len(rowTemp.strSense) > 0 Then
      strSQL = strSQL & "SENSE, "
   End If
   strSQL = strSQL & rowTemp.strType & "Code )" & CRLF()
   strSQL = strSQL & "SELECT DISTINCTROW" & CRLF() & "QCntr(" & strConstraintName & ") AS ID," & CRLF()
   strSQL = strSQL & strSelObj
   If Len(rowTemp.strRHSFld) > 0 Then
      strSQL = strSQL & rowTemp.strRecSet & "." & rowTemp.strRHSFld & " AS RHS, " & CRLF()
   End If
   If Len(rowTemp.strSense) > 0 Then
      strSQL = strSQL & "'" & rowTemp.strSense & "'" & " AS SENSE, " & CRLF()
   End If
      
   strConstraintName = strConstraintName & " AS ConstraintName" & CRLF()
   
   strSQL = strSQL & strConstraintName
   
   strSQL = strSQL & "FROM " & rowTemp.strRecSet & CRLF()
   strSQL = strSQL & ";"
   
   MakePopRowQry = strSQL

End Function

Function MakeUpdColSlnQry(colTemp As typColType) As String

Dim strUPD As String, strSQL As String

   MakeUpdColSlnQry = "FALSE"

   strUPD = "UPDATE DISTINCTROW" & CRLF()
   strSQL = strUPD & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & CRLF()
   strSQL = strSQL & "INNER JOIN " & LPM_COLUMN_TABLE_NAME & " ON " & CRLF()
   strSQL = strSQL & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & "." & colTemp.strType & "ID = " & LPM_COLUMN_TABLE_NAME & ".ColID" & CRLF()
   strSQL = strSQL & "SET" & CRLF()
   strSQL = strSQL & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & ".ACTIVITY" & " = " & LPM_COLUMN_TABLE_NAME & ".ACTIVITY," & CRLF()
   strSQL = strSQL & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & ".DJ" & " = " & LPM_COLUMN_TABLE_NAME & ".DJ," & CRLF()
   strSQL = strSQL & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & ".OBJLO" & " = " & LPM_COLUMN_TABLE_NAME & ".OBJLO," & CRLF()
   strSQL = strSQL & LPM_COL_ELEMENT_TABLE_PRE & colTemp.strTable & ".OBJUP" & " = " & LPM_COLUMN_TABLE_NAME & ".OBJUP" & CRLF()
   strSQL = strSQL & ";"

   MakeUpdColSlnQry = strSQL

End Function

Function MakeUpdRowSlnQry(rowTemp As typRowType) As String

Dim strUPD As String, strSQL As String

   MakeUpdRowSlnQry = "FALSE"

   strUPD = "UPDATE DISTINCTROW" & CRLF()
   strSQL = strUPD & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & CRLF()
   strSQL = strSQL & "INNER JOIN " & LPM_CONSTR_TABLE_NAME & " ON " & CRLF()
   strSQL = strSQL & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & "." & rowTemp.strType & "ID = " & LPM_CONSTR_TABLE_NAME & ".RowID" & CRLF()
   strSQL = strSQL & "SET" & CRLF()
   strSQL = strSQL & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & ".ACTIVITY" & " = " & LPM_CONSTR_TABLE_NAME & ".ACTIVITY," & CRLF()
   strSQL = strSQL & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & ".SHADOW" & " = " & LPM_CONSTR_TABLE_NAME & ".SHADOW," & CRLF()
   strSQL = strSQL & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & ".RHSLO" & " = " & LPM_CONSTR_TABLE_NAME & ".RHSLO," & CRLF()
   strSQL = strSQL & LPM_ROW_ELEMENT_TABLE_PRE & rowTemp.strTable & ".RHSUP" & " = " & LPM_CONSTR_TABLE_NAME & ".RHSUP" & CRLF()
   strSQL = strSQL & ";"

   MakeUpdRowSlnQry = strSQL

End Function


Function MatGen()
'//================================================================================//
'/|   FUNCTION: MatGen                                                             |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Generate the Matrix by Populating LPM_MATRIX_TABLE_NAME using the  |/
'/|             data from LPM_COLUMN_DEF_TABLE_NAME, LPM_CONSTR_DEF_TABLE_NAME,    |/
'/|             and LPM_COEFFS_DEF_TABLE_NAME                                      |/
'/|      USAGE: i= MatGen()                                                        |/
'/|         BY: Sean                                                               |/
'/|       DATE: 4/11/97                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database, qd As QueryDef, rs As Recordset, rsClone As Recordset
Dim i As Integer, j As Integer, K As Integer, intCols As Integer, intIgNull As Integer
Dim T As Date
Dim strSQL As String, strINS As String, strAltQryText As String, stemp As String
Dim rowTemp As typRowType
Dim colTemp As typColType
Dim coeffTemp As typCoeffType
Dim s As String

On Error GoTo MatGen_Err

MatGen = False
T = Now()
i = SetCntrToZero()

Set db = CurrentDb
   
Set rs = db.OpenRecordset(LPM_COEFFS_DEF_TABLE_NAME)
Set rsClone = rs.Clone()

rsClone.MoveFirst

intIgNull = True

strINS = "INSERT INTO " & LPM_MATRIX_TABLE_NAME & CRLF() & _
         "  ( COL, ColID, OBJ, LO, UP, FREE, SOS, " & _
         "    ROW, RowID, SENSE, RHS, RANGE, COE )" & CRLF()


'// C O E F F  L O O P  //
While rsClone.EOF = False
   gMatGenOK = True
   stemp = MakeMatGenQry2(rsClone!CoeffTypeID)
   
   If rsClone!UseAlternateQry = True Then
      strSQL = strINS & rsClone!AlternateQryText
   Else
      strSQL = strINS & stemp
   End If
   
   '// Currently the active switch in LPM_COEFFS_DEF_TABLE_NAME does nothing
   If gMatGenOK Then
      DoCmd.SetWarnings False
      s = rsClone!CoeffTypeID & " " & rsClone!colType & ", " & rsClone!rowType
          Debug.Print s;
          If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s
      DoCmd.RunSQL strSQL
      s = "     ... " & Format(Now() - T, "hh:nn:ss")
          Debug.Print s
          'Debug.Print strSQL  '//
          If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
      DoCmd.SetWarnings True
   End If
   rsClone.MoveNext
Wend

DoCmd.SetWarnings False
Set qd = db.QueryDefs!zqColumns
strSQL = qd.SQL
DoCmd.RunSQL strSQL
s = "C List.. " & Format(Now() - T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()

Set qd = db.QueryDefs!zqRows
strSQL = qd.SQL
DoCmd.RunSQL strSQL
s = "R List.. " & Format(Now() - T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
DoCmd.SetWarnings True
 
 
T = Now() - T
s = "MatGen....." & Format(T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()

MatGen = True
   
MatGen_Done:
  Exit Function

MatGen_Err:

Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         MatGen = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume MatGen_Done
      End If
   Case 3265
      MatGen = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query and field name exist."
      Resume MatGen_Done
   Case 3000
      MatGen = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "I think you need to check and see if the column element and row element tables exist."
      Resume MatGen_Done
    Case Else
      MatGen = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume MatGen_Done
      Resume
End Select

End Function

Function Null2BigM(AValue)
   ' Purpose: Return the value 0 if AValue is Null.
   If IsNull(AValue) Then
      Null2BigM = LPM_BIG_M
   Else
      Null2BigM = AValue
   End If
End Function
Function OutputMPS()

Dim db As Database
Set db = CurrentDb
Dim rs As Recordset
Dim strTempname As String
Dim rsClone As Recordset
Dim sSNS As String, sCol As String
Dim intRowColNameSize As Integer
Const LP_NAME = "NAME"
Const LP_OBJSENSE = "OBJSENSE"
Const LP_MAX = "MAX"
Const LP_MIN = "MIN"
Const LP_COL_HEAD = "COLUMNS"
Const LP_ROW_HEAD = "ROWS"
Const LP_RHS_HEAD = "RHS"
Const LP_BND_HEAD = "BOUNDS"
Const LP_OBJ_TYPE = "N"

intRowColNameSize = 30
Open MPSLP.strFileName For Output As #1     '//  Open file for output.

'//  Print the Problem Name Line
Print #1, LP_NAME; Spc(11); MPSLP.strProblemName

'//  Print the OBJ Row's Sense
Print #1, LP_OBJSENSE
Select Case MPSLP.intOBJSense
   Case OBJ_SENSE_MAX
      Print #1, Spc(1); LP_MAX
   Case OBJ_SENSE_MIN
      Print #1, Spc(1); LP_MIN
End Select

'//  Print the Row's Section Header
Print #1, LP_ROW_HEAD

'//  Print the Objective row
Print #1, Spc(1); LP_OBJ_TYPE; Spc(2); MPSLP.strOBJRowName     '//  Objective Row First

Set rs = db.OpenRecordset(MPSLP.strRowRS)
Set rsClone = rs.Clone()
rsClone.MoveFirst
While rsClone.EOF = False
   If rsClone!SENSE = 1 Then
      sSNS = "E"
   ElseIf rsClone!SENSE = 2 Then
      sSNS = "G"
   Else
      sSNS = "L"
   End If
   Print #1, Spc(1); sSNS; Spc(1); rsClone!ROW                         '//  Print the rest of the rows
   rsClone.MoveNext
Wend

'//  Print the Column's Section
Print #1, LP_COL_HEAD

Set rs = db.OpenRecordset(MPSLP.strCoeRS)
Set rsClone = rs.Clone()
rsClone.MoveFirst
sCol = "" 'INIT

While rsClone.EOF = False
   If sCol <> SizeName(rsClone!Col, intRowColNameSize) Then  'if this is a brand new vector then print the obj coeff.
      Print #1, Spc(3); SizeName(rsClone!Col, intRowColNameSize); SizeName(MPSLP.strOBJRowName, intRowColNameSize); Spc(5); rsClone!OBJ
      sCol = SizeName(rsClone!Col, intRowColNameSize)
   End If
   Print #1, Spc(3); SizeName(rsClone!Col, intRowColNameSize); SizeName(rsClone!ROW, intRowColNameSize); Spc(5); rsClone!COE
   rsClone.MoveNext
Wend


Close #1


End Function


Function PopElmTbls()
'//================================================================================//
'/|   FUNCTION: PopElmTbls                                                         |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Populate the tables that store individual LP vectors and rows      |/
'/|      USAGE: i= PopElmTbls()                                                    |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/21/97                                                            |/
'/|    HISTORY: 3/27/97  Added feature to generate the append queries              |/
'/|                      automatically, then run them.                             |/
'//================================================================================//

Dim db As Database, qd As QueryDef, rs As Recordset, rsClone As Recordset
Dim i As Integer, j As Integer, K As Integer, intCols As Integer, intIgNull As Integer
Dim strSQL As String, strINS As String, strColRS As String, strColTable As String
Dim strSelObj As String, strVectorName As String
Dim strColClassField As String, strColClass As String
Dim T As Date
Dim rowTemp As typRowType
Dim colTemp As typColType
Dim strColTypePrefix As String
Dim s As String

On Error GoTo PopElmTbls_Err

PopElmTbls = False
T = Now()
i = SetCntrToZero()

Set db = CurrentDb
   
'// C O L U M N S //
Set rs = db.OpenRecordset(LPM_COLUMN_DEF_TABLE_NAME)
Set rsClone = rs.Clone()

rsClone.MoveFirst

intIgNull = True

s = "----- C O L U M N S -----"
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()

While rsClone.EOF = False
   colTemp = ReadColType(rsClone!ColTypeID)
   If colTemp.intActive Then
      strSQL = MakePopColQry(colTemp)
      DoCmd.SetWarnings False
      s = colTemp.strType
          Debug.Print s
          If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
      DoCmd.RunSQL strSQL
      DoCmd.SetWarnings True
   End If
   rsClone.MoveNext
Wend

i = SetCntrToZero()

'// R O W S //
Set rs = db.OpenRecordset(LPM_CONSTR_DEF_TABLE_NAME)
Set rsClone = rs.Clone()

rsClone.MoveFirst

intIgNull = True

s = "----- R O W S -----"
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()

While rsClone.EOF = False
   rowTemp = ReadRowType(rsClone!RowTypeID)
   If rowTemp.intActive Then
      strSQL = MakePopRowQry(rowTemp)
      DoCmd.SetWarnings False
      s = rowTemp.strType
          Debug.Print s
          If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
      DoCmd.RunSQL strSQL
      DoCmd.SetWarnings True
   End If
   rsClone.MoveNext
Wend
   
T = Now() - T
s = "PopElmTbls....." & Format(T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()

PopElmTbls = True
   
PopElmTbls_Done:
  Exit Function

PopElmTbls_Err:

Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         PopElmTbls = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume PopElmTbls_Done
      End If
   Case 3265
      PopElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query and field name exist."
      Resume PopElmTbls_Done
   Case 3000
      PopElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "I think you need to check and see if the column element and row element tables exist."
      Resume PopElmTbls_Done
    Case Else
      PopElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume PopElmTbls_Done
      Resume
End Select

End Function

Function ReadCoeffType(ID As Long) As typCoeffType
'//================================================================================//
'/|   FUNCTION: ReadCoeffType                                                      |/
'/| PARAMETERS: ID, the CoeffTypeID from table LPM_COEFFS_DEF_TABLE_NAME           |/
'/|    RETURNS: CoeffType data variable of form:                                   |/
'/|                                                                                |/
'/|                lngID        As Long                                            |/
'/|                intActive    As Integer                                         |/
'/|                strType      As String                                          |/
'/|                strColType   As String                                          |/
'/|                lngColID     As Long                                            |/
'/|                strRowType   As String                                          |/
'/|                lngRowID     As Long                                            |/
'/|                strRecSet    As String                                          |/
'/|                strCoeffFld  As String                                          |/
'/|                                                                                |/
'/|    PURPOSE: Read in all the information associated with a Coeff type           |/
'/|      USAGE: CoeffTemp = ReadCoeffType(7)                                       |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/31/97                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset, rscol As Recordset, rsrow As Recordset
Dim i As Integer, j As Integer
Dim T As Date
Dim coeffTemp As typCoeffType
Dim strClassConcat As String

On Error GoTo ReadCoeffType_Err
   
   Set db = CurrentDb
   Set rs = db.OpenRecordset(LPM_COEFFS_DEF_TABLE_NAME)      '// Open the Coeff Definition Table
   Set rscol = db.OpenRecordset(LPM_COLUMN_DEF_TABLE_NAME)     '// Open the Col Definition Table
   Set rsrow = db.OpenRecordset(LPM_CONSTR_DEF_TABLE_NAME)     '// Open the Row Definition Table
   rs.MoveFirst
   rscol.MoveFirst
   rs.MoveFirst
   
   Do While rs.EOF = False                                  '// Find the Correct Coeff Type
      If rs!CoeffTypeID = ID Then                           '// Read the static Coeff data
         coeffTemp.lngID = rs!CoeffTypeID
         coeffTemp.intActive = rs!CoeffActive
         coeffTemp.strType = rs!CoeffType
         coeffTemp.strColType = rs!colType
         
         '// Get the ColTypeID
         Do While rscol.EOF = False                         '// Find the Correct Col Type
            If rscol!colType = coeffTemp.strColType Then    '// Read the Col ID
                coeffTemp.lngColID = rscol!ColTypeID
                Exit Do
            End If
            rscol.MoveNext
         Loop
         
         coeffTemp.strRowType = rs!rowType
         
         '// Get the RowTypeID
         Do While rsrow.EOF = False                         '// Find the Correct Row Type
            If rsrow!rowType = coeffTemp.strRowType Then    '// Read the Row ID
                coeffTemp.lngRowID = rsrow!RowTypeID
                Exit Do
            End If
            rsrow.MoveNext
         Loop
         
         coeffTemp.strRecSet = rs!CoeffRecSet
         coeffTemp.strCoeffFld = rs!CoeffField
         Exit Do
      End If
      rs.MoveNext
   Loop
   
ReadCoeffType = coeffTemp

ReadCoeffType_Done:
  Exit Function

ReadCoeffType_Err:
   If Err = 94 Then 'Invalid null
      Resume Next
   Else
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume ReadCoeffType_Done
      Resume
   End If
   
End Function
Function ReadColType(ID As Long) As typColType
'//================================================================================//
'/|   FUNCTION: ReadColType                                                        |/
'/| PARAMETERS: ID, the ColTypeID from table LPM_COLUMN_DEF_TABLE_NAME             |/
'/|    RETURNS: ColType data variable of form:                                     |/
'/|                                                                                |/
'/|                   lngID         As Long                                        |/
'/|                   intActive     As Integer                                     |/
'/|                   strType       As String                                      |/
'/|                   strDesc       As String                                      |/
'/|                   strTable      As String                                      |/
'/|                   strRecSet     As String                                      |/
'/|                   strPrefix     As String                                      |/
'/|                   strSOS        As String                                      |/
'/|                   intFREE       As Integer                                     |/
'/|                   strOBJFld     As String                                      |/
'/|                   strLOFld      As String                                      |/
'/|                   strUPJFld     As String                                      |/
'/|                   intClassCount As Integer                                     |/
'/|                   strClasses()  As String                                      |/
'/|                                                                                |/
'/|    PURPOSE: Read in all the information associated with a Col type             |/
'/|      USAGE: ColTemp = ReadColType(7)                                           |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/31/98                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset
Dim i As Integer, j As Integer
Dim T As Date
Dim colTemp As typColType
Dim strClassConcat As String

On Error GoTo ReadColType_Err
   
   Set db = CurrentDb
   Set rs = db.OpenRecordset(LPM_COLUMN_DEF_TABLE_NAME)     '// Open the Col Definition Table
   rs.MoveFirst
   
   Do While rs.EOF = False                               '// Find the Correct Col Type
      If rs!ColTypeID = ID Then                          '// Read the static Col data
         
         colTemp.lngID = rs!ColTypeID
         colTemp.intActive = rs!ColActive
         colTemp.strType = rs!colType
         colTemp.strDesc = rs!ColTypeDesc
         colTemp.strTable = rs!ColTypeTable
         colTemp.strRecSet = rs!ColTypeRecSet
         colTemp.strPrefix = rs!ColTypePrefix
         colTemp.strSOS = rs!SOSType
         colTemp.intFREE = rs!BNDFree
         colTemp.strOBJFld = rs!ObjField
         colTemp.strLOFld = rs!BNDLoField
         colTemp.strUPFld = rs!BNDUpField
         colTemp.intClassCount = CountColClasses(colTemp.strType)    '// Count the Classes
         
         '// ReDimension the Class Array to the proper size
         ReDim colTemp.strClasses(colTemp.intClassCount)
         
         '// Loop through the classes
         For i = 1 To colTemp.intClassCount
            colTemp.strClasses(i) = rs("C" & i)                      '// Load the array
            strClassConcat = strClassConcat & Brack(rs("C" & i))     '// Make the Concat string
         Next i
      
         With rs
            .Edit
            !ClassCount = colTemp.intClassCount  '// update the Col def table with the class count
            !ClassConcat = strClassConcat        '// update the Col def table with the class cat
            .Update
         End With
         
         Exit Do
      End If
      rs.MoveNext
   Loop
   
ReadColType = colTemp

ReadColType_Done:
  Exit Function

ReadColType_Err:
   If Err = 94 Then 'Invalid null
      Resume Next
   Else
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume ReadColType_Done
      Resume
   End If
   
End Function
Function ReadDatType(ID As Long) As typDatType
'//================================================================================//
'/|   FUNCTION: ReadDatType                                                        |/
'/| PARAMETERS: ID, the DatTypeID from table LPM_DAT_DEF_TABLE_NAME                |/
'/|    RETURNS: DatType data variable of form:                                     |/
'/|                                                                                |/
'/|                   lngID         As Long                                        |/
'/|                   intActive     As Integer                                     |/
'/|                   intMaster     As Integer                                     |/
'/|                   strType       As String                                      |/
'/|                   strDesc       As String                                      |/
'/|                   strTable      As String                                      |/
'/|                   intClassCount As Integer                                     |/
'/|                   strClasses()  As String                                      |/
'/|                                                                                |/
'/|    PURPOSE: Read in all the information associated with a Dat type             |/
'/|      USAGE: datTemp = ReadDatType(1)                                           |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/31/97                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset
Dim i As Integer, j As Integer
Dim T As Date
Dim datTemp As typDatType
Dim strClassConcat As String

On Error GoTo ReadDatType_Err
   
   Set db = CurrentDb
   Set rs = db.OpenRecordset(LPM_DAT_DEF_TABLE_NAME)     '// Open the Dat Definition Table
   rs.MoveFirst
   
   Do While rs.EOF = False                               '// Find the Correct Dat Type
      If rs!DatTypeID = ID Then                          '// Read the static Dat data
         
         datTemp.lngID = rs!DatTypeID
         datTemp.intActive = rs!DatActive
         datTemp.intMaster = rs!DatMaster
         datTemp.strType = rs!datType
         datTemp.strDesc = rs!DatTypeDesc
         datTemp.strTable = rs!DatTypeTable
         datTemp.intClassCount = CountDatClasses(datTemp.strType)   '// Count the Classes
         datTemp.intFieldCount = CountDatFields(datTemp.strType)    '// Count the Fields
         
         '// ReDimension the Class Array to the proper size
         ReDim datTemp.strClasses(datTemp.intClassCount)
         
         '// Loop through the classes
         For i = 1 To datTemp.intClassCount
            datTemp.strClasses(i) = rs("D" & i)                      '// Load the array
            strClassConcat = strClassConcat & Brack(rs("D" & i))     '// Make the Concat string
         Next i
      
         With rs
            .Edit
            !ClassCount = datTemp.intClassCount    '// update the Dat def table with the class count
            !ClassConcat = strClassConcat          '// update the Dat def table with the class cat
            !FieldCount = datTemp.intFieldCount    '// update the Dat def table with the field count
            .Update
         End With
         
         Exit Do
      End If
      rs.MoveNext
   Loop
   
ReadDatType = datTemp

ReadDatType_Done:
  Exit Function

ReadDatType_Err:
   If Err = 94 Then 'Invalid null
      Resume Next
   Else
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume ReadDatType_Done
      Resume
   End If
   
End Function
Function ReadRowType(ID As Long) As typRowType
'//================================================================================//
'/|   FUNCTION: ReadRowType                                                        |/
'/| PARAMETERS: ID, the RowTypeID from table LPM_CONSTR_DEF_TABLE_NAME             |/
'/|    RETURNS: RowType data variable of form:                                     |/
'/|                                                                                |/
'/|             RowTypeVar.lngID              As Long                              |/
'/|             RowTypeVar.intActive          As Integer                           |/
'/|             RowTypeVar.strType            As String                            |/
'/|             RowTypeVar.strDesc            As String                            |/
'/|             RowTypeVar.strTable           As String                            |/
'/|             RowTypeVar.strRecSet          As String                            |/
'/|             RowTypeVar.strPrefix          As String                            |/
'/|             RowTypeVar.strSense           As String                            |/
'/|             RowTypeVar.strRHSFld          As String                            |/
'/|             RowTypeVar.intClassCount      As Integer                           |/
'/|             RowTypeVar.strClasses()       As String                            |/
'/|                                                                                |/
'/|    PURPOSE: Read in all the information associated with a row type             |/
'/|      USAGE: rowTemp = ReadRowType(7)                                           |/
'/|         BY: Sean                                                               |/
'/|       DATE: 3/31/97                                                            |/
'/|    HISTORY:                                                                    |/
'//================================================================================//

Dim db As Database
Dim rs As Recordset
Dim i As Integer, j As Integer
Dim T As Date
Dim rowTemp As typRowType
Dim strClassConcat As String

On Error GoTo ReadRowType_Err
   
   Set db = CurrentDb
   Set rs = db.OpenRecordset(LPM_CONSTR_DEF_TABLE_NAME)     '// Open the Row Definition Table
   rs.MoveFirst
   
   Do While rs.EOF = False                               '// Find the Correct Row Type
      If rs!RowTypeID = ID Then                          '// Read the static row data
         
         rowTemp.lngID = rs!RowTypeID
         rowTemp.intActive = rs!RowActive
         rowTemp.strType = rs!rowType
         rowTemp.strDesc = rs!RowTypeDesc
         rowTemp.strTable = rs!RowTypeTable
         rowTemp.strRecSet = rs!RowTypeRecSet
         rowTemp.strPrefix = rs!RowTypePrefix
         rowTemp.strSense = rs!RowTypeSNS
         rowTemp.strRHSFld = rs!RHSField
         rowTemp.intClassCount = CountRowClasses(rowTemp.strType)    '// Count the Classes
         
         '// ReDimension the Class Array to the proper size
         ReDim rowTemp.strClasses(rowTemp.intClassCount)
         
         '// Loop through the classes
         For i = 1 To rowTemp.intClassCount
            rowTemp.strClasses(i) = rs("R" & i)                      '// Load the array
            strClassConcat = strClassConcat & Brack(rs("R" & i))     '// Make the Concat string
         Next i
      
         With rs
            .Edit
            !ClassCount = rowTemp.intClassCount  '// update the row def table with the class count
            !ClassConcat = strClassConcat        '// update the row def table with the class cat
            .Update
         End With
         
         Exit Do
      End If
      rs.MoveNext
   Loop
   
ReadRowType = rowTemp

ReadRowType_Done:
  Exit Function

ReadRowType_Err:
   If Err = 94 Then 'Invalid null
      Resume Next
   Else
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume ReadRowType_Done
      Resume
   End If
   
End Function


Function RunAll()
    Dim z As Date
    Dim zstart As Date
    Dim i As Integer
    Dim s As String
    Dim sLogFileName As String
    Dim qd As QueryDef
    Dim db As Database
    Dim rsADO As ADODB.Recordset
    Dim cnCurrentDB As ADODB.Connection
    Dim lRet         As Long 'RKP/08-12-08
    Dim timeStamp    As String 'RKP/08-12-08
    Dim filePathName As String 'RKP/08-12-08
    
    Set db = CurrentDb
    z = Now()
    zstart = Now()
    gtxtDest.Value = ""
    
    Debug.Print "Model Start" & CRLF()
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & "Model Start" & CRLF()

    Debug.Print "Append ConvProd RawUse Table" & CRLF()
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & "Append ConvProd RawUse Table" & CRLF()

    
    'TEMP TABLE FOR ITEM SALES
    Set qd = db.QueryDefs("qMTXdelCONPROD_RAWUSE")
    qd.Execute
    Set qd = db.QueryDefs("qMTXappCONPROD_RAWUSE")
    qd.Execute
    
    Debug.Print "Append Item Sales FormTable" & CRLF()
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & "Append Item Sales FormTable" & CRLF()

    'TEMP TABLE FOR ITEM SALES
    Set qd = db.QueryDefs("qMTXdelITEMSALES")
    qd.Execute
    Set qd = db.QueryDefs("qMTXappITEMSALES")
    qd.Execute
    z = Now() - z
    s = "  C_ITEMSALES TEMP TABLE...     DONE... " & Format(z, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
        
    i = InitModel()
    i = GenElmTbls()
    i = PopElmTbls()
    i = MatGen()
    i = Solve()
    i = SsvElmTbls()
    'i = TestOutputMPS() 'NOT CURRENTLY WORKING (obj)
    
    If IsTableQuery("", "tSYStblTopTen") Then DeleteTable ("tSYStblTopTen")
    Set qd = db.QueryDefs("qCAPmakSysTopTen")
    qd.Execute
    s = "  TopTen ...     DONE"
    Debug.Print s
    
    If IsTableQuery("", "tSYStblInfeasibilities") Then DeleteTable ("tSYStblInfeasibilities")
    Set qd = db.QueryDefs("qCAPmakSysInfeasibilities")
    qd.Execute
    s = "  Infeasiblities Report ...     DONE"
    Debug.Print s

    If IsTableQuery("", "tSYStblInfeasibilityCauses") Then DeleteTable ("tSYStblInfeasibilityCauses")
    Set qd = db.QueryDefs("qCAPmakSysInfeasibilityCauses")
    qd.Execute
    s = "  Infeasibility Causes ...     DONE"
    Debug.Print s

    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
    

    Dim rs As Recordset
    Set rs = db.OpenRecordset("qMTXselC_MISCINFO")
    rs.MoveFirst
    rs.Edit
    rs!LAST_RUN_DATE = Now()
    rs.Update
    rs.Close
    
    If IsTableQuery("", "trepCOLMTX") Then DeleteTable ("trepCOLMTX")
    Set qd = db.QueryDefs("qmakCOLMTX")
    qd.Execute
    
    DoCmd.SetWarnings False
''    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qrepCAPdeltagrid", _
''       "C:\OPTMODELS\PC30\Output\" & _
''       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-PMPROD-" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
''    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "trepCOLMTX", _
''       "C:\OPTMODELS\PC30\Output\" & _
''       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-COLMTX-" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True

   'RKP/08-12-08
   timeStamp = VBA.Format(VBA.Now(), "YYYYMMDD-HHMM")
   filePathName = CreateObject("Scripting.FileSystemObject").GetParentFolderName(Application.CurrentProject.Connection.Properties("Data Source Name").Value) & "\Output\" & DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-COLMTX-" & timeStamp '& ".XML.RUN"
   
   'CSV version of COLMTX
    'DoCmd.TransferText acExportDelim, "IMEXtrepCOLMTX", "trepCOLMTX", _
       "C:\OPTMODELS\PC30\Output\" & _
       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-COLMTX-" & timeStamp & ".CSV", True
       
    'RKP/08-12-08
    'Commented the above DoCmd with the following line to make use of filePathName variable
    'DoCmd.TransferText acExportDelim, "IMEXtrepCOLMTX", "trepCOLMTX", "C:\TEMP\A.CSV", True, ""
    'DoCmd.TransferText acExportDelim, "IMEXtrepCOLMTX", "trepCOLMTX", filePathName & ".CSV", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "trepCOLMTX", filePathName & ".XLS", True
    'DoCmd.TransferText acExportDelim, "IMEXtrepCOLMTX", "trepCOLMTX", "zzzz" & ".CSV", True
   'RKP/08-12-08
   'Dump the RUN Files
   'Create a serialized ADO Recordset. This XML file is an input to the Diff Compare (RunComparison) algorithm in C-OPT1.
   Call Application.CurrentProject.Connection.Execute("SELECT * FROM trepCOLMTX", lRet).Save(filePathName & ".XML.RUN", adPersistXML)

''    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "qrepCAPConvProductionXtab", _
''       "C:\OPTMODELS\PC30\Output\" & _
''       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-CONVRT-" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
''
''    'Dump the RUN Files
''    Set rsADO = Application.CurrentProject.Connection.Execute("SELECT * FROM qrepCAPdeltagrid", i)
''    rsADO.Save "C:\OPTMODELS\PC30\Output\" & _
''       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-PMPRD-" & Format(Now(), "yyyymmdd-hhmm") & ".XML.RUN", adPersistXML
''    Set rsADO = Application.CurrentProject.Connection.Execute("SELECT * FROM trepCOLMTX", i)
''    rsADO.Save "C:\OPTMODELS\PC30\Output\" & _
''       DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-COLMTX-" & Format(Now(), "yyyymmdd-hhmm") & ".XML.RUN", adPersistXML
''    'Set rsADO = Application.CurrentProject.Connection.Execute("SELECT * FROM qrepCAPConvProductionXtab", i)
''    'rsADO.Save DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-CONVRT-" & Format(Now(), "yyyymmdd-hhmm") & ".XML.RUN", adPersistXML
    
    
    
    'Set qd = db.QueryDefs("qrepmakPMcomparegrid")
    'If IsTableQuery("", "tsysMOVES-COMPARE") Then DeleteTable ("tsysMOVES-COMPARE")
    '*&*  STM 9/28/2004
    'qd.Execute
    
    's = "SELECT [qrepCOLMTX].* INTO [tsysCOLMTX-" & DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-" & _
            Format(Now(), "yyyymmdd-hhmm") & "] FROM [qrepCOLMTX];"
    'db.Execute s
    
    '//HANDLE MOVES REPORTING TABLES BASED ON RECON SWITCH
    'If DLookup("RECON", "qMTXselC_MISCINFO", "MiscInfoSetID=1") = True Then 'RECON RUN
    '  'MAKE THE BASELINE
    '  s = "SELECT [qrepCOLMTX].* INTO [tsysMOVES-BASELINE] FROM [qrepCOLMTX];"
    '  'If IsTableQuery("", "tsysMOVES-BASELINE") Then DeleteTable ("tsysMOVES-BASELINE")
    '  db.Execute s
    'Else 'NORMAL OPT RUN
    '  'MAKE THE MOVES TABLE: tsysCOLMTXMOVES
    '  Set qd = db.QueryDefs("Q09COLMTXMOVES")
    '  qd.Execute
    '  'SAVE TO COMP RUN
    '  s = "SELECT [qrepCOLMTX].* INTO [tsysMOVES-COMPARE] FROM [qrepCOLMTX];"
    '  'If IsTableQuery("", "tsysMOVES-COMPARE") Then DeleteTable ("tsysMOVES-COMPARE")
    '  db.Execute s
    'End If
    
    'okay, now the real stuff:  qrepCOLMTX
    
    'i = CompactDB("C:\OPTMODELS\PC30", "CAPDATA.MDB")
    
    DoCmd.SetWarnings True
    
    'DoCmd.OpenReport "repKI", acViewNormal   'this prints
    'DoCmd.OpenReport "repKI", acViewPreview
    
    z = Now() - zstart
    s = "Full Run..." & Format(z, "hh:nn:ss")
    Debug.Print s
    Debug.Print Now()
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
    'open file
    'log gtxtDest.Value
    sLogFileName = "C:\OPTMODELS\PC31\OUTPUT\PC31.LOG"
    If VBA.Dir(sLogFileName) <> "" Then VBA.Kill sLogFileName
    Open sLogFileName For Output As #1
    Print #1, gtxtDest.Value
    Close #1
    VBA.FileCopy sLogFileName, VBA.Left(sLogFileName, Len(sLogFileName) - 4) & "_" & _
       timeStamp & ".LOG"
End Function
Function RunAllForm(txtDest As TextBox)
    Dim i As Integer
    gFormRunFlag = True
    Set gtxtDest = txtDest
    gtxtDest.Value = ""
    gtxtDest.Value = "...OK...START... " & Now() & CRLF()
    i = RunAll()
End Function

Function Solve() As Integer

Dim sCmd As String, strSQL As String, sMPLPathAndFile As String, sSYSPath As String, sMPLModelFile As String
Dim i As Integer
Dim qd As QueryDef
Dim db As Database
Dim T As Date
Dim s As String
Dim sHostName As String

Solve = False
T = Now()
DoCmd.SetWarnings False
strSQL = "Delete * FROM " & LPM_COLUMN_TABLE_NAME_IMPORT
DoCmd.RunSQL strSQL
strSQL = "Delete * FROM " & LPM_CONSTR_TABLE_NAME_IMPORT
DoCmd.RunSQL strSQL
DoCmd.SetWarnings True

sHostName = Environ$("computername")
Debug.Print sHostName

sSYSPath = "C:\OPTMODELS\PC31"
sMPLModelFile = "BBSOPAPP.MPL"


''''sMPLPathAndFile = "C:\" & Chr(34) & "Program Files" & Chr(34) & "\MATH\MPL\50\bin\win32\MPLWIN.EXE"   'MODEL_BOX
''''sMPLPathAndFile = sSYSPath & "\SOLVE.BAT"                                                                 'MODEL_BOX

''sMPLPathAndFile = "C:\Programs\MATH\MPL\50\bin\win64\MPLWIN.EXE"   'MODEL_BOX

sMPLPathAndFile = "C:\Programs\MATH\MPL\50\bin\win32\MPLWIN.EXE"   'MODEL_BOX


If InStr(sHostName, "SMACDER") > 0 Then
   ''sMPLPathAndFile = "C:\Programs\MATH\MPL64BIT\MPLWIN42.EXE"        'STM DELL 64 Bit
   sMPLPathAndFile = "C:\Programs\MATH\MPL5.064BIT\50\bin\win64\Mplwin.exe"
   ''                 C:\Programs\MATH\MPL5.064BIT\50\bin\win64\Mplwin.exe  'OK SEAN DELL 2016.09.22
End If


sCmd = sMPLPathAndFile & " SOLVE CPLEX  " & sSYSPath & "\" & sMPLModelFile
Debug.Print "The command is hard wired to (for [" & sHostName & "] ) "
Debug.Print sCmd
               
ExecCmd (sCmd)
   
T = Now() - T
s = "Solve......" & Format(T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
T = Now()

Set db = CurrentDb
   
DoCmd.TransferText acImportDelim, "IMEX_Columns", LPM_COLUMN_TABLE_NAME_IMPORT, "C:\OPTMODELS\PC31\SGO_C.TXT"
DoCmd.TransferText acImportDelim, "IMEX_Row1", LPM_CONSTR_TABLE_NAME_IMPORT, "C:\OPTMODELS\PC31\SGO_R1.TXT"
DoCmd.TransferText acImportDelim, "IMEX_Row1", LPM_CONSTR_TABLE_NAME_IMPORT, "C:\OPTMODELS\PC31\SGO_R2.TXT"
DoCmd.TransferText acImportDelim, "IMEX_Row1", LPM_CONSTR_TABLE_NAME_IMPORT, "C:\OPTMODELS\PC31\SGO_R3.TXT"
   
DoCmd.SetWarnings False
Set qd = db.QueryDefs!qupdColumns
strSQL = qd.SQL
DoCmd.RunSQL strSQL

Set qd = db.QueryDefs!qupdRows
strSQL = qd.SQL
DoCmd.RunSQL strSQL
DoCmd.SetWarnings True

T = Now() - T
s = "Import Soln...." & Format(T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()
   
Solve = True

Solve_Done:
  Exit Function

Solve_Err:
   If Err = 94 Then 'Invalid null
      Resume Next
   Else
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume Solve_Done
      Resume
   End If

End Function


Function TestOutputMPS() As Integer

Dim i As Integer
TestOutputMPS = False

MPSLP.strFileName = GetPathName(GetCurrentMDBwPath()) & "FOO.MPS"
MPSLP.strProblemName = "FOO"
MPSLP.intOBJSense = OBJ_SENSE_MAX
MPSLP.strOBJRowName = "PROFIT"
MPSLP.strRowRS = LPM_CONSTR_TABLE_NAME
MPSLP.strColRS = LPM_COLUMN_TABLE_NAME
MPSLP.strCoeRS = LPM_MATRIX_TABLE_NAME

i = OutputMPS()
TestOutputMPS = i

End Function

Function TestPopColQry(j As Long) As Integer

Dim foo As String
Dim i As Integer
Dim col9 As typColType

TestPopColQry = False
foo = ""

col9 = ReadColType(j)

Debug.Print "ID     =" & col9.lngID
Debug.Print "ON     =" & col9.intActive
Debug.Print "TYPE   =" & col9.strType
Debug.Print "DESC   =" & col9.strDesc
Debug.Print "TABLE  =" & col9.strTable
Debug.Print "RECSET =" & col9.strRecSet
Debug.Print "PRE    =" & col9.strPrefix
Debug.Print "SOS    =" & col9.strSOS
Debug.Print "FREE   =" & col9.intFREE
Debug.Print "OBJ    =" & col9.strOBJFld
Debug.Print "LO     =" & col9.strLOFld
Debug.Print "UP     =" & col9.strUPFld
Debug.Print "#CLASS =" & col9.intClassCount

For i = 1 To col9.intClassCount
   foo = foo & Brack(col9.strClasses(i))
Next i

Debug.Print "CLASS$ =" & foo
Debug.Print "POPSQL =" & MakePopColQry(col9)


i = CreateColTable(col9)
         
TestPopColQry = True

End Function

Function SsvElmTbls()
'//================================================================================//
'/|   FUNCTION: SsvElmTbls                                                         |/
'/| PARAMETERS: -NONE-                                                             |/
'/|    RETURNS: True on Success and False by default or Failure                    |/
'/|    PURPOSE: Store solution values for the individual LP vectors and rows       |/
'/|      USAGE: i= SsvElmTbls()                                                    |/
'/|         BY: Sean                                                               |/
'/|       DATE: 4/15/97                                                            |/
'/|    HISTORY: 4/15/97  Added feature to generate the append queries              |/
'//================================================================================//

Dim db As Database, qd As QueryDef, rs As Recordset, rsClone As Recordset
Dim i As Integer, j As Integer, K As Integer, intCols As Integer, intIgNull As Integer
Dim strSQL As String, strINS As String, strColRS As String, strColTable As String
Dim strSelObj As String, strVectorName As String
Dim strColClassField As String, strColClass As String
Dim T As Date
Dim rowTemp As typRowType
Dim colTemp As typColType
Dim strColTypePrefix As String
Dim s As String

On Error GoTo SsvElmTbls_Err

SsvElmTbls = False
T = Now()


Set db = CurrentDb
   
'// C O L U M N S //
Set rs = db.OpenRecordset(LPM_COLUMN_DEF_TABLE_NAME)
Set rsClone = rs.Clone()

rsClone.MoveFirst

intIgNull = True

Debug.Print "----- C O L U M N S -----"

While rsClone.EOF = False
   colTemp = ReadColType(rsClone!ColTypeID)
   If colTemp.intActive Then
      strSQL = MakeUpdColSlnQry(colTemp)
      DoCmd.SetWarnings False
      Debug.Print colTemp.strType
      DoCmd.RunSQL strSQL
      DoCmd.SetWarnings True
   End If
   rsClone.MoveNext
Wend


'// R O W S //
Set rs = db.OpenRecordset(LPM_CONSTR_DEF_TABLE_NAME)
Set rsClone = rs.Clone()

rsClone.MoveFirst

intIgNull = True

Debug.Print "----- R O W S -----"

While rsClone.EOF = False
   rowTemp = ReadRowType(rsClone!RowTypeID)
   If rowTemp.intActive Then
      strSQL = MakeUpdRowSlnQry(rowTemp)
      DoCmd.SetWarnings False
      Debug.Print rowTemp.strType
      DoCmd.RunSQL strSQL
      DoCmd.SetWarnings True
   End If
   rsClone.MoveNext
Wend
   
T = Now() - T
s = "SsvElmTbls....." & Format(T, "hh:nn:ss")
    Debug.Print s
    If gFormRunFlag Then gtxtDest.Value = gtxtDest.Value & s & CRLF()


SsvElmTbls = True
   
SsvElmTbls_Done:
  Exit Function

SsvElmTbls_Err:

Select Case Err
   Case 94 'Invalid Null
      If intIgNull Then
         Resume Next
      Else
         SsvElmTbls = False
         MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
         Resume SsvElmTbls_Done
      End If
   Case 3265
      SsvElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "Check to make sure that the query and field name exist."
      Resume SsvElmTbls_Done
   Case 3000
      SsvElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description & CRLF() & "I think you need to check and see if the column element and row element tables exist."
      Resume SsvElmTbls_Done
    Case Else
      SsvElmTbls = False
      MsgBox "An unexpected error (" & Err.Number & ") occurred:  " & Err.Description
      Resume SsvElmTbls_Done
      Resume
End Select

End Function
Function test12() As Integer

Dim s As String
Dim i As Integer
test12 = False
'txtDest.Value = txtDest.Value & "All Right Y'all"
    Dim rsADO As ADODB.Recordset
    Set rsADO = Application.CurrentProject.Connection.Execute("SELECT * FROM qrepCAPdeltagrid", i)
    rsADO.Save DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-PMPRD-" & Format(Now(), "yyyymmdd-hhmm") & ".XML.RUN", adPersistXML
    Set rsADO = Application.CurrentProject.Connection.Execute("SELECT * FROM trepCOLMTX", i)
    rsADO.Save DLookup("RUN_NAME", "tSYStblRunName", "ID=1") & "-COLMTX-" & Format(Now(), "yyyymmdd-hhmm") & ".XML.RUN", adPersistXML
test12 = True
End Function
Function ExportTablesToExcel()
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPtblMillProduction", _
       "C:\OPTMODELS\PC31\Output\" & "tCAPtblMillProduction_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPtblMillProduction2", _
       "C:\OPTMODELS\PC31\Output\" & "tCAPtblMillProduction2_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPtblMillXfers", _
       "C:\OPTMODELS\PC31\Output\" & "tCAPtblMillXfers_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPtblDemand", _
       "C:\OPTMODELS\PC31\Output\" & "tCAPtblDemand_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPlkpScheduleGrade", _
       "C:\OPTMODELS\PC31\Output\" & "ScheduleGrade_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    DoCmd.TransferSpreadsheet acExport, acSpreadsheetTypeExcel9, "tCAPlkpItem", _
       "C:\OPTMODELS\PC31\Output\" & "Item_" & Format(Now(), "yyyymmdd-hhmm") & ".XLS", True
    End Function
Function TestCoeffRead(j As Long) As Integer

Dim foo As String
Dim i As Integer
Dim coeff As typCoeffType

TestCoeffRead = False
foo = ""

coeff = ReadCoeffType(j)

Debug.Print "ID     =" & coeff.lngID
Debug.Print "ON     =" & coeff.intActive
Debug.Print "TYPE   =" & coeff.strType
Debug.Print "COLTYPE=" & coeff.strColType
Debug.Print "COLID  =" & coeff.lngColID
Debug.Print "ROWTYPE=" & coeff.strRowType
Debug.Print "ROWID  =" & coeff.lngRowID
Debug.Print "RECSET =" & coeff.strRecSet
Debug.Print "FIELD  =" & coeff.strCoeffFld

TestCoeffRead = True

End Function
Function TestColRead(j As Long) As Integer

Dim foo As String
Dim i As Integer
Dim s As String
Dim col9 As typColType

TestColRead = False
foo = ""

col9 = ReadColType(j)

Debug.Print "ID     =" & col9.lngID
Debug.Print "ON     =" & col9.intActive
Debug.Print "TYPE   =" & col9.strType
Debug.Print "DESC   =" & col9.strDesc
Debug.Print "TABLE  =" & col9.strTable
Debug.Print "RECSET =" & col9.strRecSet
Debug.Print "PRE    =" & col9.strPrefix
Debug.Print "SOS    =" & col9.strSOS
Debug.Print "FREE   =" & col9.intFREE
Debug.Print "OBJ    =" & col9.strOBJFld
Debug.Print "LO     =" & col9.strLOFld
Debug.Print "UP     =" & col9.strUPFld
Debug.Print "#CLASS =" & col9.intClassCount

For i = 1 To col9.intClassCount
   foo = foo & Brack(col9.strClasses(i))
Next i

Debug.Print "CLASS$ =" & foo
         
'i = CreateColTable(col9)
's = MakeUpdColSlnQry(col9)

Debug.Print s
         
TestColRead = True

End Function

Function TestData(j As Long) As Integer

Dim foo As String
Dim i As Integer
Dim dat1 As typDatType

TestData = False
foo = ""

dat1 = ReadDatType(j)

Debug.Print "ID     =" & dat1.lngID
Debug.Print "ON     =" & dat1.intActive
Debug.Print "TYPE   =" & dat1.strType
Debug.Print "DESC   =" & dat1.strDesc
Debug.Print "TABLE  =" & dat1.strTable
Debug.Print "#CLASS =" & dat1.intClassCount
Debug.Print "#FIELD =" & dat1.intFieldCount

For i = 1 To dat1.intClassCount
   foo = foo & Brack(dat1.strClasses(i))
Next i

Debug.Print "CLASS$ =" & foo
         
i = CreateDatTable(dat1)
         
TestData = i

End Function


Function TestRows() As Integer

Dim foo As String
Dim i As Integer
Dim row7 As typRowType

TestRows = False
foo = ""

row7 = ReadRowType(1)

Debug.Print "ID     =" & row7.lngID
Debug.Print "ON     =" & row7.intActive
Debug.Print "TYPE   =" & row7.strType
Debug.Print "DESC   =" & row7.strDesc
Debug.Print "TABLE  =" & row7.strTable
Debug.Print "RECSET =" & row7.strRecSet
Debug.Print "PRE    =" & row7.strPrefix
Debug.Print "SENSE  =" & row7.strSense
Debug.Print "RHS    =" & row7.strRHSFld
Debug.Print "#CLASS =" & row7.intClassCount

For i = 1 To row7.intClassCount
   foo = foo & Brack(row7.strClasses(i))
Next i

Debug.Print "CLASS$ =" & foo
         
'i = CreateRowTable(row7)
foo = MakeUpdRowSlnQry(row7)
Debug.Print foo
         
TestRows = True

End Function
Function UnBrack(strIn As String) As String
   
   If (Left(strIn, 1) = "[") Then
      strIn = Right(strIn, Len(strIn) - 1)
   End If
   
   If (Right(strIn, 1) = "]") Then
      strIn = Left(strIn, Len(strIn) - 1)
   End If
   
   UnBrack = strIn

End Function






