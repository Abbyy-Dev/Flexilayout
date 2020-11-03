Dim date, dayMonth, year

'28-05-20
date = me.Field("invoicedate").Text

dayMonth = Mid(date,1,6)
year = Mid(date,7,2)


if me.Field("invoicedate").Text <> "" then

		me.Field("invoicedate").Text = dayMonth + "20" + year
		
end if