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