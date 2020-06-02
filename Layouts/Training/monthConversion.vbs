3 mai 2020

4 mars
4 juin
4 août

5 avril

7 janvier
7 février
7 octobre
7 juillet

8 novembre
8 décembre

9 septembre


1 mars 2020
11 mars 2020

Dim cdate, day1, day2, month2_3, month2_4, month2_5, month2_7, month2_8, month2_9, year2_3, year2_4, year2_5, year2_7, year2_8, year2_9, month1_3, month1_4, month1_5, month1_7, month1_8, month1_9, year1_3, year1_4, year1_5, year1_7, year1_8, year1_9


cdate = me.Field("orderdatebuyer").Text
day1 = Mid(cdate,1,1)
day2 = Mid(cdate,1,2)

'--------- 1 digit
'2. mai 2020
month1_3 = Mid(cdate,4,3)
year1_3 = Mid(cdate,8,4)

'4. mars 2020
month1_4 = Mid(cdate,4,4)
year1_4 = Mid(cdate,9,4)

'5. avril 2020
month1_5 = Mid(cdate,4,5)
year1_5 = Mid(cdate,10,4)

'7. janvier 2020
month1_7 = Mid(cdate,4,7)
year1_7 = Mid(cdate,12,4)

'8. novembre 2020
month1_8 = Mid(cdate,4,8)
year1_8 = Mid(cdate,13,4)

'9. septembre 2020
month1_9 = Mid(cdate,4,9)
year1_9 = Mid(cdate,14,4)

'--------- 2 digits
'20. mai 2020
month2_3 = Mid(cdate,5,3)
year2_3 = Mid(cdate,9,4)

'14. mars 2020
month2_4 = Mid(cdate,5,4)
year2_4 = Mid(cdate,10,4)

'25. avril 2020
month2_5 = Mid(cdate,5,5)
year2_5 = Mid(cdate,11,4)

'20. janvier 2020
month2_7 = Mid(cdate,5,7)
year2_7 = Mid(cdate,13,4)

'18. novembre 2020
month2_8 = Mid(cdate,5,8)
year2_8 = Mid(cdate,14,4)

'19. septembre 2020
month2_9 = Mid(cdate,5,9)
year2_9 = Mid(cdate,15,4)


if me.Field("orderdatebuyer").Text <> "" and Len(day1) = 1 then
   if Len(month1_3) = 3 then
        if month1_3 = "mai" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "05" + "-" + year1_3
		end if
	
	elseif Len(month1_4) = 4 then
        if month1_4 = "mars" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "03" + "-" + year1_4
		elseif month1_4 = "juin" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "06" + "-" + year1_4
		elseif month1_4 = "août" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "08" + "-" + year1_4
		end if
	
	elseif Len(month1_5) = 5 then
        if month1_5 = "avril" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "04" + "-" + year1_5
		end if
	
	elseif Len(month1_7) = 7 then
        if month1_7 = "janvier" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "01" + "-" + year1_7
        elseif month1_7 = "février" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "02" + "-" + year1_7
		elseif month1_7 = "juillet" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "07" + "-" + year1_7
        elseif month1_7 = "octobre" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "10" + "-" + year1_7
        end if	
		
	elseif Len(month1_8) = 8 then
        if month1_8 = "novembre" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "11" + "-" + year1_8
		elseif month1_ = "décembre" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "12" + "-" + year1_8
		end if
		
	elseif Len(month1_9) = 9 then
        if month1_8 = "septembre" then
        me.Field("orderdatebuyer").Text = day1 + "-" + "09" + "-" + year1_9
		end if
		
	end if
	
elseif me.Field("orderdatebuyer").Text <> "" and Len(day2) = 2 then
    if Len(month2_3) = 3 then
        if month2_3 = "mai" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "05" + "-" + year2_3
		end if
	
	elseif Len(month2_4) = 4 then
        if month2_4 = "mars" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "03" + "-" + year2_4
		elseif month2_4 = "juin" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "06" + "-" + year2_4
		elseif month2_4 = "août" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "08" + "-" + year2_4
		end if
	
	elseif Len(month2_5) = 5 then
        if month2_5 = "avril" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "04" + "-" + year2_5
		end if
	
	elseif Len(month2_7) = 7 then
        if month2_7 = "janvier" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "01" + "-" + year2_7
        elseif month2_7 = "février" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "02" + "-" + year2_7
		elseif month2_7 = "juillet" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "07" + "-" + year2_7
        elseif month2_7 = "octobre" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "10" + "-" + year2_7
        end if
		
	elseif Len(month2_8) = 8 then
        if month2_8 = "novembre" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "11" + "-" + year2_8
		elseif month2_8 = "décembre" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "12" + "-" + year2_8
		end if
		
	elseif Len(month2_9) = 9 then
        if month2_9 = "septembre" then
        me.Field("orderdatebuyer").Text = day2 + "-" + "09" + "-" + year2_9
		end if
		
    end if
	

end if