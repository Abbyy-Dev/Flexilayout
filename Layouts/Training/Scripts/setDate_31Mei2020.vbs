Dim date,year, year2, smonth, smonth2 ,day, day2

'1 mei 2020
'31 mei 2020
date = me.Field("invoicedate").Text
year = Mid(date,7,4) '1 mei 2020
year2 = Mid(date,8,4) '31 mei 2020
smonth = Mid(date,3,3) '1 mei 2020
smonth2 = Mid(date,4,3) '31 mei 2020
day = Mid(date,1,1) '1 mei 2020
day2 = Mid(date,1,2)'31 mei 2020


if me.Field("invoicedate").Text <> ""  then 
  'jan
  if smonth = "jan" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "01" + "-"+ year
  elseif smonth2 = "jan" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "01" + "-"+ year2
    
  'feb
  elseif smonth = "feb" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "02" + "-"+ year  
  elseif smonth2 = "feb" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "02" + "-"+ year2
    
  'mrt
  elseif smonth = "mrt" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "03" + "-"+ year  
  elseif smonth2 = "mrt" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "03" + "-"+ year2
  
  'apr
  elseif smonth = "apr" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "04" + "-"+ year
  elseif smonth2 = "apr" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "04" + "-"+ year2
    
  'mei
  elseif smonth = "mei" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "05" + "-"+ year
  elseif smonth2 = "mei" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "05" + "-"+ year2

    
  'juni
  elseif smonth = "juni" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "06" + "-"+ year
  elseif smonth2 = "juni" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "06" + "-"+ year2
    
   'juli
  elseif smonth = "juli" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "07" + "-"+ year
  elseif smonth2 = "juli" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "07" + "-"+ year2
    
   'aug
  elseif smonth = "aug" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "08" + "-"+ year
  elseif smonth2 = "aug" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "08" + "-"+ year2
    
   'sep
  elseif smonth = "sep" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "09" + "-"+ year
  elseif smonth2 = "sep" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "09" + "-"+ year2
    
   'okt
  elseif smonth = "okt" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "10" + "-"+ year
  elseif smonth2 = "okt" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "10" + "-"+ year2  
    
   'nov
  elseif smonth = "nov" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "11" + "-"+ year
  elseif smonth2 = "nov" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "11" + "-"+ year2  
    
   'dec
  elseif smonth = "dec" and Len(day) = 1 then
     me.Field("invoicedate").Text = "0" + day + "-" + "12" + "-"+ year
  elseif smonth2 = "dec" and Len(day2) = 2 then
     me.Field("invoicedate").Text = day2 + "-" + "12" + "-"+ year2  
    
  end if

else
     me.Field("invoicedate").Text = date
  
end if