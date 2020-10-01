=================== !!! LATEST from dutch date!!!  ============ Updated 2020-09-14 

Dim cdate, year, sep1, x, y, month

'31 April 2020
cdate = LCase(me.Field("invoicedate").Text)
year = Right(cdate,4)
sep1 = CStr(InStr(1,cdate," "))
x = CInt(InStr(1,cdate," ")+1) 'always start after the first sep1
y = CInt(InStrRev(cdate," ")-x) 
month = Mid(cdate,x,y)

if cdate <> "" and sep1 = "3" then
    if month = "january" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "01" & "-" & year
    elseif month = "february" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "02" & "-" & year
    elseif month = "maarch" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "03" & "-" & year
    elseif month = "april" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "04" & "-" & year
    elseif month = "may" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "05" & "-" & year
    elseif month = "june" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "06" & "-" & year
    elseif month = "july" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "07" & "-" & year
    elseif month = "august" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "08" & "-" & year
    elseif month = "september" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "09" & "-" & year
    elseif month = "october" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "10" & "-" & year
    elseif month = "november" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "11" & "-" & year
    elseif month = "december" then
       me.Field("invoicedate").Text = Mid(cdate,1,2) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
elseif cdate <> "" and sep1 = "2" then
    if month = "january" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "01" & "-" & year
    elseif month = "february" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "02" & "-" & year
    elseif month = "march" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "03" & "-" & year
    elseif month = "april" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "04" & "-" & year
    elseif month = "may" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "05" & "-" & year
    elseif month = "june" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "06" & "-" & year
    elseif month = "july" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "07" & "-" & year
    elseif month = "august" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "08" & "-" & year
    elseif month = "september" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "09" & "-" & year
    elseif month = "october" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "10" & "-" & year
    elseif month = "november" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "11" & "-" & year
    elseif month = "december" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,1) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
end if

=======================================


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