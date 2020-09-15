=============== 23-7-2020 =================

Dim duedate, day, month, year, day2, month2, day3, month3, day4, month4


duedate = Replace(me.Field("duedate").Text," ","") '23-7-2020


'6-11-2020 = 9
day = Mid(duedate,1,1)
month = Mid(duedate,3,2)

'23-7-2020 = 9
day2 = Mid(duedate,1,2)
month2 = Mid(duedate,4,1)

'7-6-2020 = 8
day3 = Mid(duedate,1,1)
month3 = Mid(duedate,3,1)

'10-10-2020 = 10
day4 = Mid(duedate,4,2)
month4 = Mid(duedate,1,2)

year = Right(duedate,4)



if me.Field("duedate").Text <> "" and Len(duedate) = 9 then
 
    if Len(day2) = 2 then
  
    me.Field("duedate").Text = day2 + "-" + "0" + month2 + "-" + year
   
 else 
   
    me.Field("duedate").Text = "0" + day + "-" + month + "-" + year
    
    end if
   
elseif me.Field("duedate").Text <> "" and Len(duedate) = 8 then

 me.Field("duedate").Text =  "0" + day3 + "-" + "0" + month3 + "-" + year
 
elseif me.Field("duedate").Text <> "" and Len(duedate) = 10 then

 me.Field("duedate").Text  = day4 + "-" + month4 + "-" + year

end if