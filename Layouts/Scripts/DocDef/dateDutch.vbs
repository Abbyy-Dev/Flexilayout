Dim cdate, year, day, x, y, month

'11 juni 2020
cdate = me.Field("invoicedate").Text
year = Right(cdate,4)
day = CStr(InStr(1,cdate," ")-1)
x = CInt(InStr(1,cdate," ")+1) '3
y = CInt(InStr(cdate,year)-5) '13-4 = 9
month = Mid(cdate,x,y)

if cdate <> "" and day = "2" then
    if month = "januari" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "februari" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "maart" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "april" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mei" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juni" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "juli" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "augustus" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "september" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "oktober" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "november" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "december" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
elseif cdate <> "" and day = "1" then
    if month = "januari" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "februari" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "maart" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "april" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mei" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juni" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "juli" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "augustus" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "september" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "oktober" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "november" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "december" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
end if




=================== 3 digit month ===============

Dim cdate, year, day, x, y, month

'8-okt-2019
cdate = me.Field("invoicedate").Text
year = Right(cdate,4) '2019
day = CStr(InStr(1,cdate,"-")-1) '3-1 = 2
x = CInt(InStr(1,cdate,"-")+1) '3
y = CInt(InStr(cdate,year)-5) '3
month = Mid(cdate,x,y)'okt

if me.Field("invoicedate").Text <> "" and day = "2" then
    if month = "jan" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "febr" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "mrt" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "apr" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mei" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juni" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "aug" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août " then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "sep" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "okt" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "nov" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "dec" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
elseif me.Field("invoicedate").Text <> "" and day = "1" then
    if month = "jan" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "febr" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "mrt" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "apr" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mei" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juni" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "aug" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août " then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "sep" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "okt" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "nov" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "dec" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
end if






====================  setting DueDate  ================================

Dim duedate, cdate, ndate, nduedate, day, day2, month, x, y, z, year

cdate = me.Field("invoicedate").Text '06-07-2020
ndate = Mid(cdate,4,2) & "/" & Mid(cdate,1,2) & "/" & Right(cdate,4) '06/07/2020
duedate = me.Field("duedate").Text '20

nduedate = CStr(DateAdd("d",duedate,ndate)) '06/27/2020

'm/d/yyyy


'05/11/2019
'5/-12-2019
year = Right(nduedate,4)
month = CStr(Instr(1,nduedate,"/")-1) '1 char
x = CStr(InStr(nduedate,year)-4) '2 char 
y = CStr(Instr(1,nduedate,"/")+1) 'pos(/) 2+1 = pos 3
z = CStr(InStr(nduedate,year)-5) '1 char 5-12-2019
day = Mid(nduedate,y,x) '2
day2 = Mid(nduedate,y,z) '1

if duedate <> "30" and Len(month) = 1  and Len(day) = 2 then 'dd/m/yyyy
    me.Field("duedate").Text = day & "-" & "0" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(month) = 1 and Len(day2) = 1 then 'd/mm/yyyy
    me.Field("duedate").Text = "0" & day2 & "-" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(month) = 1 and Len(day) = 2 then 'dd/mm/yyyy
    me.Field("duedate").Text = "0" & day2 & "-" & "0" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(month) = 1 and Len(day2) = 1 then 'd/m/yyyy
    me.Field("duedate").Text = "0" & day2 & "-" & "0" & Mid(nduedate,1,month) & "-" & year
end if


=========================='25-2-2019======================================

Dim duedate, cdate, ndate, nduedate, day, day2, month, z, year

cdate = me.Field("invoicedate").Text '25-2-2019
ndate = Mid(cdate,1,2) & "/" & Mid(cdate,4,2) & "/" & Right(cdate,4) '06/07/2020
duedate = me.Field("duedate").Text '20
nduedate = CStr(DateAdd("d",duedate,ndate)) '06/27/2020

'm/d/yyyy
'05/11/2019
'5/-12-2019

year = Right(nduedate,4)
month = CStr(Instr(1,nduedate,"/")-1) '1 char
y = CStr(Instr(1,nduedate,"/")+1) 'pos(/) 2+1 = pos 3
day = Mid(nduedate,y,1)
day2 = Mid(nduedate,y,2)

if duedate <> "30" and Len(day2) = 1 and Len(month) = 1 then 'dd/m/yyyy
    me.Field("duedate").Text = day2 & "-" & "0" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(day) = 1 and Len(month) = 2 then 'd/mm/yyyy
    me.Field("duedate").Text = "0" & day & "-" & "0" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(day2) = 2 and Len(month) = 2 then 'dd/mm/yyyy
    me.Field("duedate").Text = "0" & day2 & "-" & Mid(nduedate,1,month) & "-" & year
elseif ddate <> "30" and Len(day) = 1 and Len(month) = 1 then 'd/m/yyyy
    me.Field("duedate").Text = "0" & day & "-" & "0" & Mid(nduedate,1,month) & "-" & year
end if

========================== date 07-06-20 ===============


Dim cdate, year, dm

'17-06-20
cdate = me.Field("invoicedate").Text
year = Right(cdate,2) '20
dm = Mid(cdate,1,6) '17-06-


if me.Field("invoicedate").Text <> "" then
    me.Field("invoicedate").Text = dm & "20" & year
end if




================== 9 maart 2020 =================


Dim cdate, year, day, x, y, z, month, month2

'9 maart 2020
cdate = me.Field("invoicedate").Text
year = Right(cdate,4)
day = CStr(InStr(1,cdate," ")-1)
x = CInt(InStr(1,cdate," ")+1) '3
y = CInt(InStr(cdate,year)-5) '13-4 = 9
z = CInt(InStr(cdate,year)-4)
month = Mid(cdate,x,y)
month2 = Mid(cdate,x,z)

if cdate <> "" and day = "2" then
    if month = "januari" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "februari" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "maart" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "april" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mei" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juni" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "juli" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "augustus" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "september" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "oktober" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "november" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "december" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if
elseif cdate <> "" and day = "1" then
    if month2 = "januari" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month2 = "februari" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month2 = "maart" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month2 = "april" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month2 = "mei" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month2 = "juni" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month2 = "juli" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month2 = "augustus" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month2 = "september" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month2 = "oktober" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month2 = "november" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month2 = "december" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("invoicedate").Text = cdate
    end if  
end if
