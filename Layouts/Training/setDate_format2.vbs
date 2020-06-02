Dim date,year,month,day,year2,month2,day2

date = me.Field("orderdatebuyer").Text
year = "20" + Mid(date,7,2)
year2 = "20" + Mid(date,6,2)
month = Mid(date,4,2)
month2 = Mid(date,3,2)
day = Mid(date,1,2)
day2 = "0" + Mid(date,1,1)


if me.Field("orderdatebuyer").Text <> "" and Len(date) = 7 then
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
