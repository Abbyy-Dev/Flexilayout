=============== 23-7-2020 ========= UPDATED 2020-09-23 ========

Dim invoicedate, day, month, year, day2, month2, day3, month3, day4, month4, day1_sep, day2_sep


invoicedate = Replace(me.Field("invoicedate").Text," ","") '23-7-2020
sep1 = CStr(InStr(1,invoicedate,"-")) '2

'6-11-2020 = 9
day1_sep = CStr(InStr(1,invoicedate,"-")) 'return position 2
day = Mid(invoicedate,1,1)
month = Mid(invoicedate,3,2)

'23-7-2020 = 9
day2_sep = CStr(InStr(1,invoicedate,"-")) 'return position 3
day2 = Mid(invoicedate,1,2)
month2 = Mid(invoicedate,4,1)

'7-6-2020 = 8
day3 = Mid(invoicedate,1,1)
month3 = Mid(invoicedate,3,1)

'10-10-2020 = 10
day4 = Mid(invoicedate,1,2)
month4 = Mid(invoicedate,4,2)

year = Right(invoicedate,4)



if me.Field("invoicedate").Text <> "" and Len(invoicedate) = 9 then
 
    if day1_sep = 3 then
  
    me.Field("invoicedate").Text = day2 + "-" + "0" + month2 + "-" + year
   
 else 
   
    me.Field("invoicedate").Text = "0" + day + "-" + month + "-" + year
    
    end if
   
elseif me.Field("invoicedate").Text <> "" and Len(invoicedate) = 8 then

 me.Field("invoicedate").Text =  "0" + day3 + "-" + "0" + month3 + "-" + year
 
elseif me.Field("invoicedate").Text <> "" and Len(invoicedate) = 10 then

 me.Field("invoicedate").Text  = day4 + "-" + month4 + "-" + year
 
 

end if


=============== 2 DIGIT YEAR === 13-5-20 ========= UPDATED 2020-09-29 ========

Dim invoicedate, day, month, year, day2, month2, day3, month3, day4, month4, day1_sep, day2_sep


invoicedate = Replace(me.Field("invoicedate").Text," ","") '23-7-2020
sep1 = CStr(InStr(1,invoicedate,"-")) '2

'6-11-20 = 7
day1_sep = CStr(InStr(1,invoicedate,"-")) 'return position 2
day = Mid(invoicedate,1,1)
month = Mid(invoicedate,3,2)

'23-7-20 = 7
day2_sep = CStr(InStr(1,invoicedate,"-")) 'return position 3
day2 = Mid(invoicedate,1,2)
month2 = Mid(invoicedate,4,1)

'7-6-20 = 6
day3 = Mid(invoicedate,1,1)
month3 = Mid(invoicedate,3,1)

'10-10-20 = 8
day4 = Mid(invoicedate,1,2)
month4 = Mid(invoicedate,4,2)

year = Right(invoicedate,2)



if me.Field("invoicedate").Text <> "" and Len(invoicedate) = 7 then
 
    if day1_sep = 3 then
  
    me.Field("invoicedate").Text = day2 + "-" + "0" + month2 + "-" + "20" + year
   
 else 
   
    me.Field("invoicedate").Text = "0" + day + "-" + month + "-" + "20" + year
    
    end if
   
elseif me.Field("invoicedate").Text <> "" and Len(invoicedate) = 6 then

 me.Field("invoicedate").Text =  "0" + day3 + "-" + "0" + month3 + "-" + "20" + year
 
elseif me.Field("invoicedate").Text <> "" and Len(invoicedate) = 8 then

 me.Field("invoicedate").Text  = day4 + "-" + month4 + "-" + "20" + year

end if