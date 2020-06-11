Dim date, cdate, day, month, year, day2, month2, year2, day3, month3, year3, day4, month4, year4

'20 December 2019
date = Replace(me.Field("invoicedate").Text," ","-")

'mm-dd-yyyy
cdate = Replace(DateValue(date),"/","-")

'mm-dd-yyyy = 10
'm-dd-yyyy = 9
'mm-d-yyyy = 9
'm-d-yyyy = 8

'mm-dd-yyyy
day = Mid(cdate,4,2)
month = Mid(cdate,1,2)
year = Mid(cdate,7,4)

'mm-d-yyyy
day2 = Mid(cdate,4,1)
month2 = Mid(cdate,1,2)
year2 = Mid(cdate,6,4)

'm-dd-yyyy
'1-19-2020
day3 = Mid(cdate,3,2)
month3 = Mid(cdate,1,1)
year3 = Mid(cdate,6,4)

'm-d-yyyy
day4 = Mid(cdate,3,1)
month4 = Mid(cdate,1,1)
year4 = Mid(cdate,5,4)


if me.Field("invoicedate").Text <> "" and Len(cdate) = 10 then
   me.Field("invoicedate").Text = day + "-" + month + "-" + year

elseif me.Field("invoicedate").Text <> "" and Len(cdate) = 9 then
    if Len(day3) = 2 then
        me.Field("invoicedate").Text = day3 + "-" + "0" + month3 + "-" + year3
    else
        me.Field("invoicedate").Text = "0" + day2 + "-" + month2 + "-" + year2
    end if
elseif me.Field("invoicedate").Text <> "" and Len(cdate) = 8 then
    me.Field("invoicedate").Text = "0" + day4 + "-" + "0" + month4 + "-" + year4
    
end if