Dim date,year,year2,smonth,day

date = me.Field("orderdatebuyer").Text
year = "20" + Mid(date,8,2)
year2 = "20" + Mid(date,9,2)
smonth = Mid(date,4,3)
day = Mid(date,1,2)


if me.Field("orderdatebuyer").Text <> "" and smonth = "Jan" then
  me.Field("orderdatebuyer").Text = day + "-" + "01" + "-"+ year

elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Feb" then
  me.Field("orderdatebuyer").Text = day + "-" + "02" + "-"+ year

elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Mar" then
  me.Field("orderdatebuyer").Text = day + "-" + "03" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Apr" then
  me.Field("orderdatebuyer").Text = day + "-" + "04" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "May" then
  me.Field("orderdatebuyer").Text = day + "-" + "05" + "-"+ year

elseif me.Field("orderdatebuyer").Text <> "" and smonth = "M a" then
  me.Field("orderdatebuyer").Text = day + "-" + "05" + "-" + year2
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Jun" then
  me.Field("orderdatebuyer").Text = day + "-" + "06" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Jul" then
  me.Field("orderdatebuyer").Text = day + "-" + "07" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Aug" then
  me.Field("orderdatebuyer").Text = day + "-" + "08" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Sep" then
  me.Field("orderdatebuyer").Text = day + "-" + "09" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Oct" then
  me.Field("orderdatebuyer").Text = day + "-" + "10" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Nov" then
  me.Field("orderdatebuyer").Text = day + "-" + "11" + "-"+ year
  
elseif me.Field("orderdatebuyer").Text <> "" and smonth = "Dec" then
  me.Field("orderdatebuyer").Text = day + "-" + "12" + "-"+ year
end if