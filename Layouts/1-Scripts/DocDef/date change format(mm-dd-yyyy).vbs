============ format 05-11-2019 to ->>> 11-05-2019============ 
Dim cdate, day, month, year


cdate = me.Field("invoicedate").Text '05-11-2019

month = Mid(cdate,1,2)
day = Mid(cdate,4,2)
year  = Mid(cdate,7,4)

if me.Field("invoicedate").Text <> "" then
    me.Field("invoicedate").Text = day & "-" & month & "-" & year '11-05-2019
	
else 
	me.Field("invoicedate").Text = cdate
	
end if