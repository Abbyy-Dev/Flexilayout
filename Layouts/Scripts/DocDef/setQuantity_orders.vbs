Dim tsubtotal, tprice, ttsubtotal, ttprice, nQuantity

tprice = Replace(me.Field("item_price").Text,".","")
tsubtotal = Replace(me.Field("subtotal").Text,".","")

ttprice = Replace(me.Field("item_price").Text,",",".")
ttsubtotal = Replace(me.Field("subtotal").Text,",",".")


nQuantity = ttsubtotal / ttprice

if me.Field("items_ordered").Text = "" then

    me.Field("items_ordered").Text = nQuantity
    
end if