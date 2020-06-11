Dim date, day, day2, year, nyear,
maart, nmarch, march

date = me.Field("invoicedate").Text

'8 maart 2020
day = Mid(date,1,1)

'18 maart 2020
day2 = Mid(date,1,2)

'2020'
year = Mid(date,InStr(date,"2020"),4) '10

'march
march = InStr(date,"maart") '3
nmarch = Mid(date,march,5) 'maart



if me.Field("invoicedate").Text <> "" then

    me.Field("invoicedate").Text = "03" + year

end if