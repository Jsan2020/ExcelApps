Attribute VB_Name = "Module5"
Sub Send_Email()
Attribute Send_Email.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Send_Email Macro
'

'
   ' Select the range of cells on the active worksheet.
   ActiveSheet.Range("P1:U300").Select
   
   ' Show the envelope on the ActiveWorkbook.
   ActiveWorkbook.EnvelopeVisible = True
   
   ' Set the optional introduction field thats adds
   ' some header text to the email body. It also sets
   ' the To and Subject lines. Finally the message
   ' is sent.
   With ActiveSheet.MailEnvelope
      .Introduction = "New Cardlock Table 1 Prices"
      .Item.To = "nwoodbury@carsonteam.com"
      .Item.Subject = "New Cardlock Table 1 Prices"
      .Item.Send
   End With
End Sub

