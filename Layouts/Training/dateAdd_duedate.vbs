Dim duedate, cdate, nduedate, day, month, year, day2, month2, year2, day3, month3, year3, day4, month4, year4


duedate = me.Field("duedate").Text
cdate = me.Field("invoicedate").Text

nduedate = Replace(DateAdd("d",duedate,cdate),"/","-") '5/8/2020

'5-8-2020 = 8 chars
day = Mid(nduedate,3,1)
month = Mid(nduedate,1,1)
year = Mid(nduedate,5,4)

'5-18-2020 = 9 chars
day2 = Mid(nduedate,3,2)
month2 = Mid(nduedate,1,1)
year2 = Mid(nduedate,6,4)

'10-8-2020 = 9 chars
day3 = Mid(nduedate,4,1)
month3 = Mid(nduedate,1,2)
year3 = Mid(nduedate,6,4)


'10-18-2020 = 10 chars
day4 = Mid(nduedate,4,2)
month4 = Mid(nduedate,1,2)
year4 = Mid(nduedate,6,4)

if me.Field("duedate").Text <> "" and Len(nduedate) = 8 then

    me.Field("duedate").Text = "0" + day + "-" + "0" + month + "-" + year
	
elseif me.Field("duedate").Text <> "" and Len(nduedate) = 9 then
	
	    if Len(day2) = 2 then
    
	        me.Field("duedate").Text = day2 + "-" + "0" + month2 + "-" + year2
	
	    else
	
	        me.Field("duedate").Text = "0" + day3 + "-" + month3 + "-" + year3
	    end if

elseif me.Field("duedate").Text <> "" and Len(nduedate) = 10 then

	me.Field("duedate").Text = day4 + month4 + "-" + year4
	    
end if

=================== script for testing

Dim duedate, cdate, nduedate

duedate = me.Field("duedate").Text
cdate = me.Field("invoicedate").Text

nduedate = DateAdd("d",duedate,cdate)


============= 11-6-2019 ============

Dim duedate, cdate, nduedate, day, month, year


duedate = me.Field("duedate").Text
cdate = me.Field("invoicedate").Text

nduedate = Replace(DateAdd("d",duedate,cdate),"/","-") '11-6-2019


'11-6-2019
day = Mid(nduedate,4,1)
month = Mid(nduedate,1,2)
year = Right(nduedate,7,4)

if me.Field("duedate").Text <> "" then

	me.Field("duedate").Text = day + "-" + month + "-" + year
	    
end if