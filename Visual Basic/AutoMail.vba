' Remember to change the year when appropriate
' But it seems like it'll save more effort with a manual edit than having to type the year in the input box each time
Public Sub GenBRM()
  Dim newMail As Object
  Dim inDate() As String
  inDate = Split(InputBox("Input date (mm.dd)"), ".")
  Set newMail = Application.CreateItem(olMailItem)
  newMail.Subject = inDate(0) & "/" & inDate(1) & "/2026 Bank Sheets"
  With newMail
    .To = "[semicolon-separated To emails]"
    .CC = "[semicolon-separated CC emails]"
    .HTMLBody = "<span style='font-size: 11pt'><p>Hello,</p>" & _
"<p>Here are the bank sheets from " & inDate(0) & "." & inDate(1) & ".2026</p>" & _
"<p>Please let me know if you have any questions.</p>" & _
"<p>Thank you,</p>" & _
"<p>[Name]<br>" & _
"[Company]</p></span>"
  End With
  newMail.Display
End Sub
