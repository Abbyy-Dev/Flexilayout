=============== LATEST!!! 8-07-2020 =================

Dim invoicedate, day, month, year


invoicedate = me.Field("invoicedate").Text '8-07-2020

'6-11-2020 = 9
day = Mid(invoicedate,1,1)
month = Mid(invoicedate,3,2)
year = Right(invoicedate,4)

if me.Field("invoicedate").Text <> "" and Len(invoicedate) = 9 then
 
	me.Field("invoicedate").Text  = "0" + day + "-" + month + "-" + year
	
else

	me.Field("invoicedate").Text  = invoicedate

end if