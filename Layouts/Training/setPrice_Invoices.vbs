Dim tsubtotal, titems, ttsubtotal, ttitems, nPrice

titems = Replace(me.Field("items_ordered").Text,".","")
tsubtotal = Replace(me.Field("line_amount").Text,".","")

ttitems = Replace(titems,",",".")
ttsubtotal = Replace(tsubtotal,",",".")

nPrice = ttsubtotal / ttitems

if me.Field("item_price").Text = "" then

    me.Field("item_price").Text = nPrice
    
end if