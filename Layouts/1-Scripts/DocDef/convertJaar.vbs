Dim tyear, year, tprice, price

tyear = CInt(me.Field("unit").Text)
tprice = me.Field("item_price").Text
year = tyear/12
price = tprice * year

if me.Field("unit").Text <> "" then

    me.Field("item_price").Text = year & "year" & " = " & price
    
end if