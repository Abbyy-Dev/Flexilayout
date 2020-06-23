nedapfrance

janvier - /
fevrier - /

mars - x '24. mars 2020
avril - x '2. avril 2020

octobre - /
novembre - x

septembre - x '17. septembre 2020


=============================

Dim date, day, day2, year, may, march, june, august, april, january, february, october, july, november, december, september

date = Replace(me.Field("orderdatebuyer").Text,".","")

'------1 digit
day = Mid(date,1,1)
'------2 digit
day2 = Mid(date,1,2)
'------2020
year = Right(date,2) '20


may = Mid(date,InStr(date,"mai"),3) '3

march = Mid(date,InStr(date,"mars"),4) '4
june = Mid(date,InStr(date,"juin"),4) '4
august = Mid(date,InStr(date,"août"),4) '4

april = Mid(date,InStr(date,"avril"),5) '5

january = Mid(date,InStr(date,"janvier"),7) '14 janvier 2020
february = Mid(date,InStr(date,"février"),7) '14 février 2020
october = Mid(date,InStr(date,"octobre"),7) '7
july = Mid(date,InStr(date,"juillet"),7) '7

november = Mid(date,InStr(date,"novembre"),8) '8
december = Mid(date,InStr(date,"décembre"),8) '8

september = Mid(date,InStr(date,"septembre"),9) '9

'============================= d-mm-yyyy
if me.Field("orderdatebuyer").Text <> "" and day = 1 then
     if may = "mei" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "05" + "-" + "20" + year
		  
     elseif march = "mars" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "03" + "-" + "20" + year	  
		  
     elseif june = "juin" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "06" + "-" + "20" + year
		  	  
     elseif august = "août" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "08" + "-" + "20" + year	  
		  
     elseif april = "avril" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "04" + "-" + "20" + year
		  	  
     elseif january = "janvier" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "01" + "-" + "20" + year
		  	  	  
     elseif february = "février" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "02" + "-" + "20" + year
		  	  	  
     elseif october = "octobre" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "10" + "-" + "20" + year
		  	  	  
     elseif july = "juillet" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "07" + "-" + "20" + year
		  	  	  	  
     elseif november = "novembre" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "11" + "-" + "20" + year
		  	  	  
     elseif december = "décembre" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "12" + "-" + "20" + year
		  	  	  
     elseif september = "septembre" then
          me.Field("orderdatebuyer").Text = "0" + day + "-" + "09" + "-" + "20" + year
		  
     end if
'============================= dd-mm-yyyy
elseif me.Field("orderdatebuyer").Text <> "" and day2 = 2 then
     if may = "mei" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "05" + "-" + "20" + year
		  
     elseif march = "mars" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "03" + "-" + "20" + year	  
		  
     elseif june = "juin" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "06" + "-" + "20" + year
		  	  
     elseif august = "août" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "08" + "-" + "20" + year	  
		  
     elseif april = "avril" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "04" + "-" + "20" + year
		  	  
     elseif january = "janvier" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "01" + "-" + "20" + year
		  	  	  
     elseif february = "février" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "02" + "-" + "20" + year
		  	  	  
     elseif october = "octobre" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "10" + "-" + "20" + year
		  	  	  
     elseif july = "juillet" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "07" + "-" + "20" + year
		  	  	  	  
     elseif november = "novembre" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "11" + "-" + "20" + year
		  	  	  
     elseif december = "décembre" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "12" + "-" + "20" + year
		  	  	  
     elseif september = "septembre" then
          me.Field("orderdatebuyer").Text = day2 + "-" + "09" + "-" + "20" + year
		  
     end if
	 
end if