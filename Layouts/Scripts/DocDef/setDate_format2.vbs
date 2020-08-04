Dim date,year,month,day,year2,month2,day2

date = me.Field("orderdatebuyer").Text
year = "20" + Mid(date,7,2)
year2 = "20" + Mid(date,6,2)
month = Mid(date,4,2)
month2 = Mid(date,3,2)
day = Mid(date,1,2)
day2 = "0" + Mid(date,1,1)


if me.Field("orderdatebuyer").Text() and Len(date) = 7 then
  me.Field("orderdatebuyer").Text = day2 + "-" + month2 + "-"+ year2
  
elseif me.Field("orderdatebuyer").Text <> "" and Len(date) > 7 then
  me.Field("orderdatebuyer").Text = day + "-" + month + "-"+ year  
  
end if


-------------------------
Dim date, day, smonth ,y

date = me.Field("orderdatebuyer").Text
day = Mid(date,1,2)
smonth = CStr(Month(date))
y = CStr(Year(date))

if me.Field("orderdatebuyer").Text <> "" and Len(smonth) = 1 then
  me.Field("orderdatebuyer").Text = day + "-" + "0"+ smonth + "-" + y
end if

'1 9 - 1 2 -2019 - autosignal
Dim cdate, cmonth1,cmonth2, cday1, cday2, cyear

cdate = me.Field("orderdatebuyer").Text
cday1 = CStr(Mid(cdate,1,1))
cday2 = CStr(Mid(cdate,3,1))
cmonth1 = CStr(Mid(cdate,7,1))
cmonth2 = CStr(Mid(cdate,9,1))
cyear = CStr(Mid(cdate,12,4))

if me.Field("orderdatebuyer").Text <> "" then
   me.Field("orderdatebuyer").Text = cday1 + cday2 + "-" + cmonth1 + cmonth2 + "-" + cyear
end if



=======date format using period =========

Dim cdate, ndate

cdate = me.Field("orderdatebuyer").Text
ndate = Replace(cdate,".","-")

if me.Field("orderdatebuyer").Text <> "" then
me.Field("orderdatebuyer").Text = ndate
end if




===one digit date format=======

Dim date, day, month ,year

'7-6-2020
date = me.Field("duedate").Text
day = Mid(date,1,1) '7
month =  Mid(date,3,1) '9
year = Mid(date,5,4) '2020

'd-mm-yyyy'
if me.Field("duedate").Text <> "" and Len(day) = 1 then
  me.Field("duedate").Text = "0" + day + "-" + month + "-" + year

'dd-m-yyyy
elseif me.Field("duedate").Text <> "" and Len(month) = 1 then
  me.Field("duedate").Text = day + "-" + "0" + month + "-" + year

'd-m-yyyy
elseif me.Field("duedate").Text <> "" and Len(month) = 1 and Len(day) = 1 then
  me.Field("duedate").Text = "0" + day + "-" + "0" + month + "-" + year

'dd-mm-yyyy
else
  me.Field("duedate").Text = date

end if



==========email==============


Dim email

email = me.Field("supplieremail").Text

if me.Field("supplieremail").Text <> "" then 
    me.Field("supplieremail").Text = "info@dtvconsultants.nl"
end if

