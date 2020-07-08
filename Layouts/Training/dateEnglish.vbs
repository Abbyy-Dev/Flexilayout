Dim date, cdate, day, month, year, day2, month2, day3, month3, day4, month4


cdate = me.Field("invoicedate").Text '28 April 2020

date = FormatDateTime(cdate,2) '4/28/2020

'11/6/2019 = 9
day = Mid(date,4,1)
month = Mid(date,1,2)


'1/29/2019 = 9
day2 = Mid(date,3,2)
month2 = Mid(date,1,1)

'5/9/2020 = 8
day3 = Mid(date,3,1)
month3 = Mid(date,1,1)

'10/10/2020 = 10
day4 = Mid(date,4,2)
month4 = Mid(date,1,2)

year = Right(date,4)


if me.Field("invoicedate").Text <> "" and Len(date) = 9 then
 
    if Len(day2) = 2 then
  
    me.Field("invoicedate").Text = day2 + "-" + "0" + month2 + "-" + year
   
 else 
   
    me.Field("invoicedate").Text = "0" + day + "-" + month + "-" + year
    
    end if
   
elseif me.Field("invoicedate").Text <> "" and Len(date) = 8 then

 me.Field("invoicedate").Text =  "0" + day3 + "-" + "0" + month3 + "-" + year
 
elseif me.Field("invoicedate").Text <> "" and Len(date) = 10 then

 me.Field("invoicedate").Text  = day4 + "-" + month4 + "-" + year

end if