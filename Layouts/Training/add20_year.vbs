Dim date, year, month, day 

'12.05.20
date = me.Field("orderdatebuyer").Text
year = "20" + Mid(date,7,2)
month = Mid(date,4,2)
day = Mid(date,1,2)

if me.Field("orderdatebuyer").Text <> "" then
    me.Field("orderdatebuyer").Text = day + "-" + month + "-" + year
end if