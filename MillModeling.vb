Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Web
Imports System.Configuration
Imports System.Reflection
Imports ADODB
Module MillModeling
    'C# to VB.NET Translation Tool:
    'http://www.kamalpatel.net/ConvertCSharp2VB.aspx
    'http://authors.aspalliance.com/aldotnet/examples/translate.aspx

    Public Function wslite_01_ConvertDataTableToRecordset(ByVal table As DataTable) As ADODB.Recordset
        Dim adoRS As New ADODB.Recordset
        Dim fieldList(table.Columns.Count) As Object
        Dim i, j As Integer
        Dim fieldName As String
        Dim adoType As ADODB.DataTypeEnum
        Dim adoSize As Integer




        'Create a new (disconnected) Recordset


        'Create an array for the names of the colums to use
        'later with the AddNew method of the Recordset.
        'object [] fieldList = new object[table.Columns.Count];
        'fieldList = New Object(table.Columns.Count)

        '  // Loop through all of the columns in the DataTable
        '  // and define each as a new field in the Recordset;
        '  // also add the name to the fieldList array.
        '  for (int i=0; i < table.Columns.Count; i++)
        '  {
        '    // get the field name, or create one if necessary
        '    string fieldName = table.Columns[i].ColumnName;
        '    if ( fieldName == null || fieldName.Length == 0 )
        '      fieldName = string.Format("Column{0}", i);
        '    fieldList[i] = fieldName;
        '
        '    // Lookup the field's equivalent ADO type and maximum size
        '    ADODB.DataTypeEnum adoType;
        '    int adoSize;
        '    GetADOType(table.Columns[i].DataType, out adoType, out adoSize);
        '
        '    // add the field to the ADO Recordset
        '    recordset.Fields.Append(fieldName, adoType, adoSize,
        '      ADODB.FieldAttributeEnum.adFldIsNullable, null);
        '
        '    // These generic values appear to work well with any size
        '    // decimal/numeric; setting them for other types does not
        '    // cause harm, but they are necessary for decimals.
        '    recordset.Fields[i].Precision = 0;
        '    recordset.Fields[i].NumericScale = 25; //the maximum
        '  } //for

        '  // Loop through all of the columns in the DataTable
        '  // and define each as a new field in the Recordset;
        '  // also add the name to the fieldList array.
        For i = 0 To table.Columns.Count - 1
            '    // get the field name, or create one if necessary
            fieldName = table.Columns(i).ColumnName
            If fieldName Is Nothing Or fieldName.Length = 0 Then
                fieldName = String.Format("Columns(0)", i)
            End If
            fieldList(i) = fieldName

            '    // Lookup the field's equivalent ADO type and maximum size
            'GetADOType(table.Columns(i).DataType, adoType, adoSize)

            '    // add the field to the ADO Recordset
            adoRS.Fields.Append(fieldName, adoType, adoSize, ADODB.FieldAttributeEnum.adFldIsNullable, Nothing)

            '    // These generic values appear to work well with any size
            '    // decimal/numeric; setting them for other types does not
            '    // cause harm, but they are necessary for decimals.
            adoRS.Fields(i).Precision = 0
            adoRS.Fields(i).NumericScale = 25
        Next i

        '
        '  // Now that the fields are defined, open the Recordset so we may
        '  // add the data. Notice that because C# does not work like VB with
        '  // optional parameters, we use the special Missing object which
        '  // is defined in System.Reflection.
        '  recordset.Open(Missing.Value, Missing.Value,
        '    ADODB.CursorTypeEnum.adOpenUnspecified,
        '    ADODB.LockTypeEnum.adLockUnspecified, -1);
        adoRS.Open(Nothing, Nothing, ADODB.CursorTypeEnum.adOpenUnspecified, ADODB.LockTypeEnum.adLockUnspecified, -1)
        '
        '  for (int i=0; i < table.Rows.Count; i++ )
        For i = 0 To table.Rows.Count
            '  {
            '    // get the current record's values into an array
            '    object [] values = table.Rows[i].ItemArray;
            Dim values() As Object = table.Rows(i).ItemArray
            '
            '    // ADO does not recognize GUID, or unique identifiers,
            '    // unless they are wrapped in curly braces, we check for
            '    // any occurences of them here and convert them to a
            '    // string in the format "{00000000-0000-0000-0000-000000000000}"
            '    for (int j=0; j < values.Length; j++ )
            For j = 0 To values.Length
                '      if ( values[j] is System.Guid )
                If TypeOf values(j) Is System.Guid Then
                    '        values[j] = '{' + values[i].ToString() + '}';
                    values(j) = "[" + values(i).GetType.ToString() + "]"
                    '
                    '    // add the current record to the ADO Recordset
                    '    recordset.AddNew(fieldList, values);
                    adoRS.AddNew(fieldList, values)
                    '  }
                End If
            Next

        Next

        '
        '  // Done - Recordset complete!
        '  return recordset;
        Return adoRS
        '}

    End Function

    Public Sub wslite_02_GetADOType(ByVal dotNetType As Type, ByRef adoType As ADODB.DataTypeEnum, ByRef adoSize As Integer)
        '// Based on the .NET data type, return the appropriate ADO data type and
        '// size to be used when defining the Recordset.

        'Initialize
        adoType = ADODB.DataTypeEnum.adEmpty
        adoSize = -1

        If dotNetType Is System.Type.GetType("System.Boolean") Then
            adoType = DataTypeEnum.adBoolean
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Byte") Then
            adoType = DataTypeEnum.adUnsignedTinyInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Byte()") Then
            adoType = DataTypeEnum.adBinary
            adoSize = 32767
        ElseIf dotNetType Is System.Type.GetType("System.DateTime") Then
            adoType = DataTypeEnum.adDate
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Decimal") Then
            adoType = DataTypeEnum.adDecimal
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Double") Then
            adoType = DataTypeEnum.adDouble
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Guid") Then
            adoType = DataTypeEnum.adGUID
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Int16") Then
            adoType = DataTypeEnum.adSmallInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Int32") Then
            adoType = DataTypeEnum.adInteger
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Int64") Then
            adoType = DataTypeEnum.adBigInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.SByte") Then
            adoType = DataTypeEnum.adTinyInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.Single") Then
            adoType = DataTypeEnum.adSingle
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.String") Then
            adoType = DataTypeEnum.adVarWChar
            adoSize = 32767
        ElseIf dotNetType Is System.Type.GetType("System.UInt16") Then
            adoType = DataTypeEnum.adUnsignedSmallInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.UInt32") Then
            adoType = DataTypeEnum.adUnsignedInt
            adoSize = -1
        ElseIf dotNetType Is System.Type.GetType("System.UInt64") Then
            adoType = DataTypeEnum.adUnsignedBigInt
            adoSize = -1
        End If
    End Sub 'GetADOType

    Public Function wslite_03_ConvertRecordsetToXmlText(ByVal recordset As ADODB.Recordset) As String
        Dim stream As ADODB.Stream = New ADODB.Stream
        '  ADODB.Stream stream = new ADODB.Stream();
        '
        '  stream.Open(Missing.Value, ADODB.ConnectModeEnum.adModeUnknown,
        '    ADODB.StreamOpenOptionsEnum.adOpenStreamUnspecified, null, null);
        '  recordset.Save(stream, ADODB.PersistFormatEnum.adPersistXML);
        '  stream.Position = 0;
        '
        '  return stream.ReadText(-1);

        stream.Open(Missing.Value, ADODB.ConnectModeEnum.adModeUnknown, ADODB.StreamOpenOptionsEnum.adOpenStreamUnspecified, "", "")
        recordset.Save(stream, ADODB.PersistFormatEnum.adPersistXML)
        stream.Position = 0

        Return stream.ReadText(-1)

    End Function

End Module
