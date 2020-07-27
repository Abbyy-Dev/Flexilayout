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