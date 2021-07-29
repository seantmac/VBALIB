Public Function SendMailforTRIMBOOK()
'you are probably looking for SendAlerts("ALL", "ALL", "103")
   Dim db As DAO.Database
   Dim RS As DAO.Recordset
   Dim RSusers As DAO.Recordset
   Dim fld As DAO.Field
   Dim strMsg As String
   Dim strSub As String
   Dim strTo As String
   Dim i As Integer
   Dim ispc As Integer
   Dim ilen As Integer
   
         
   Set db = CurrentDb
   ispc = 10
   
   DoCmd.SetWarnings False
   
   '*&*
   
   Set RSusers = db.OpenRecordset("qryTrimBookEmailUsers", dbOpenDynaset)
   
   RSusers.MoveLast
   RSusers.MoveFirst
   
   If Not RSusers.EOF Then
      strUserName = RSusers!SENDTOUSR
   
   Set RS = db.OpenRecordset("qryTrimBookEmail", dbOpenDynaset)
   
   strTo = "nicholas.willis@ipaper.com"  'SENDTOUSRID  'nwill17@ipaper.com
   strSub = "Trimmed Orders Vs. Bookings - " & VBA.Now()

   If Not RS.BOF Then
      strMsg = "Please check to see why orders are trimmed without bookings."
      strMsg = strMsg & vbCrLf & vbCrLf
      For Each fld In RS.Fields
         Select Case fld.OrdinalPosition
         Case 0, 1, 6, 8, 15
            ispc = 10
         Case 2
            ispc = 10
         Case 7, 9, 10, 11, 12, 13, 14
            ispc = 10
         Case 3, 4
            ispc = 10
         Case Else
            ispc = 26
         End Select
         ilen = Len(fld.Name)
         If ispc > ilen Then ispc = ispc - ilen
         strMsg = strMsg & fld.Name & Space(ispc)
      Next fld
      strMsg = strMsg & vbCrLf & vbCrLf
      Do Until RS.EOF
         For Each fld In RS.Fields
            Select Case fld.OrdinalPosition
            Case 0, 1, 6, 8, 15
               ispc = 10
            Case 2
               ispc = 10
            Case 7, 9, 10, 11, 12, 13, 14
               ispc = 10
            Case 3, 4
               ispc = 10
            Case Else
               ispc = 26
            End Select
            ilen = Null2Zero(Len(Trim(fld.value)))
            If ispc > ilen Then ispc = ispc - ilen
            strMsg = strMsg & Trim(fld.value) & Space(ispc)
         Next fld
         strMsg = strMsg & vbCrLf
         RS.MoveNext
      Loop
    
      i = SendMail(strMsg, strTo, "", strSub)
      
      'SendMail(ByVal emailBody As String, ByVal emailTo As String, ByVal emailCC As String, Optional emailSubject)
      'DoCmd.SendObject acSendNoObject, "qryALERTNewStudent", acFormatPDF, "Sean <mailto:doano10@macdart.com>; STM <mailto:stm@macdart.com>", , , _
                           strSub, strMsg, False
      'docmd.SendObject acSendNoObject
      
      '*&*
      
   End If

myExit:
   DoCmd.SetWarnings True
   If Not RS Is Nothing Then RS.Close: Set RS = Nothing
   'DoCmd.CloseDatabase
   Exit Function

myError:
   Select Case Err.Number
       Case 99999
           ' trap specific errors
       Case Else
           MsgBox "Exception Nr. " & Err.Number & ". " & Err.Description
           Resume myExit
   End Select

End Function