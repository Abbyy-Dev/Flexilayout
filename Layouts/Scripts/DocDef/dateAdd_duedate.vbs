=================== !!! LATEST !!! 29-10-2019  ==========================================================

Dim ddate, idate, ndate, cdate, year, month, y, day, day2


idate = me.Field("invoicedate").Text '17-06-2020
ndate = Mid(idate,4,2) & "/" & Mid(idate,1,2) & "/" & Right(idate,4) '06/17/2020
ddate = CInt(me.Field("duedate").Text) '30

cdate = CStr(DateAdd("d",ddate,ndate)) '7/17/2020
year = Right(cdate,4) '2020'
month = CStr(Instr(1,cdate,"/")-1) '7
y = CStr(Instr(1,cdate,"/")+1) 'pos3
day = Mid(cdate,y,1) '7/7/2020
day2 = Mid(cdate,y,2) '7/17/2020

'17-07-2020 dd-mm-yyyy
if me.Field("duedate").Text <> "" and Len(day2) = 2 and Len(month) = 1 then 'dd-m-yyyy
    me.Field("duedate").Text = day2 & "-" & "0" & Mid(cdate,1,month) & "-" & year
elseif me.Field("duedate").Text <> "" and Len(day) = 1 and Len(month) = 1 then 'd-m-yyyy
    me.Field("duedate").Text = "0" & day & "-" & "0" & Mid(cdate,1,month) & "-" & year
elseif me.Field("duedate").Text <> "" and Len(day) = 1 and Len(month) = 2 then 'd-mm-yyyy
    me.Field("duedate").Text = "0" & day & "-" & Mid(cdate,1,month) & "-" & year
elseif me.Field("duedate").Text <> "" and Len(day2) = 2 and Len(month) = 2 then 'dd-mm-yyyy
    me.Field("duedate").Text = day & "-" & Mid(cdate,1,month) & "-" & year
end if

============================   5/8/2020   ==================================


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



=================== script for testing =========================================================================

Dim duedate, cdate, nduedate

duedate = me.Field("duedate").Text
cdate = me.Field("invoicedate").Text

nduedate = DateAdd("d",duedate,cdate)


============= 11-6-2019 =========================================================================

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


============= 01-11-2019 =========================================================================

Dim ddate, idate, ndate, cdate, year, month, x, y, day
'11-06-2020
idate = me.Field("invoicedate").Text
ndate = Mid(idate,4,2) & "/" & Mid(idate,1,2) & "/" & Right(idate,4)
ddate = CInt(me.Field("duedate").Text) '14

'6/25/2020 - 9
cdate = CStr(DateAdd("d",ddate,ndate))
year = Right(cdate,4)
month = CStr(Instr(1,cdate,"/")-1)
x = CStr(InStr(cdate,year)-4)
y = CStr(Instr(1,cdate,"/")+1)
day = Mid(cdate,y,x)


if ddate <> "30" and month = "1" and Len(day) = 2 then
    me.Field("duedate").Text = day & "-" & "0" & Mid(cdate,1,month) & "-" & year
elseif ddate <> "30" and month = "1" and Len(day) = 1 then
    me.Field("duedate").Text = "0" & day & "-" & Mid(cdate,1,month) & "-" & year
else
    me.Field("duedate").Text = day & "-" & Mid(cdate,1,month) & "-" & year
end if
