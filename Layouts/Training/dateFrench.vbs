Dim cdate, year, day, x, y, month

'2. novembre 2020
'15. janvier 2020
cdate = me.Field("orderdatebuyer").Text
year = Right(cdate,4)
day = CStr(InStrRev(cdate,".",-1)-1)
x = CInt(InStr(cdate,".")+2) '4
y = CInt(InStr(cdate,year)-5) '13-4 = 9
month = Mid(cdate,x,y)

if cdate <> "" and day = "2" then
    if month = "janvier " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "février " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "mars " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "avril " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mai " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juin " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "juillet " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "septembre " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "octobre " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "novembre " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "décembre " then
       me.Field("orderdatebuyer").Text = Mid(cdate,1,day) & "-" & "12" & "-" & year
    else 
       me.Field("orderdatebuyer").Text = cdate
    end if
elseif cdate <> "" and day = "1" then
    if month = "janvier" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "01" & "-" & year
    elseif month = "février" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "02" & "-" & year
    elseif month = "mars" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "03" & "-" & year
    elseif month = "avril" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "04" & "-" & year
    elseif month = "mai" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "05" & "-" & year
    elseif month = "juin" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "06" & "-" & year
    elseif month = "juillet" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "07" & "-" & year
    elseif month = "août" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "08" & "-" & year
    elseif month = "septembre" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "09" & "-" & year
    elseif month = "octobre" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "10" & "-" & year
    elseif month = "novembre" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "11" & "-" & year
    elseif month = "décembre" then
       me.Field("orderdatebuyer").Text = "0" & Mid(cdate,1,day) & "-" & "12" & "-" & year
    else
       me.Field("orderdatebuyer").Text = cdate
    end if
end if


