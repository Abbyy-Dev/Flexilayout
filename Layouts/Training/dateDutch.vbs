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
    elseif month = "augustus" then
       me.Field("invoicedate").Text = Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août " then
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
    elseif month = "augustus" then
       me.Field("invoicedate").Text = "0" & Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août " then
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


'setting DueDate
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


if ddate <> "" and month = "1" and Len(day) = 2 then
    me.Field("duedate").Text = day & "-" & "0" & Mid(cdate,1,month) & "-" & year
elseif ddate <> "" and month = "1" and Len(day) = 1 then
    me.Field("duedate").Text = "0" & day & "-" & Mid(cdate,1,month) & "-" & year
else
    me.Field("duedate").Text = day & "-" & Mid(cdate,1,month) & "-" & year
end if