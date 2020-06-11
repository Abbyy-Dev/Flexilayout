Dim date, day, month, year, day2, month2, year2, day3, month3, year3, day4, month4, year4

'26-5-2020
date = me.Field("invoicedate").Text

'd-m-yyyy = 8
day = Mid(date,1,1)
month = Mid(date,3,1)
year = Mid(date,5,4)


'dd-m-yyyy = 9
day2 = Mid(date,1,2)
month2 = Mid(date,4,1)
year2 = Mid(date,6,4)

'd-mm-yyyy = 9
day3 = Mid(date,1,1)
month3 = Mid(date,3,2)
year3 = Mid(date,6,4)


'dd-mm-yyyy = 10
day4 = Mid(date,1,2)
month4 = Mid(date,4,2)
year4 = Mid(date,7,4)


if me.Field("invoicedate").Text <> "" and Len(date) = 8 then

       me.Field("invoicedate").Text = "0" + day + "-" + "0" + month + "-" + year

elseif me.Field("invoicedate").Text <> "" and Len(date) = 9 then

    if Len(day2) = 2 then
        me.Field("invoicedate").Text = day2 + "-" + "0" + month2 + "-" + year2
    else
        me.Field("invoicedate").Text = "0" + day3 + "-" + month3 + "-" + year3
    end if

elseif me.Field("invoicedate").Text <> "" and Len(date) = 10 then

    me.Field("invoicedate").Text = day4 + "-" + month4 + "-" + year4
    
end if