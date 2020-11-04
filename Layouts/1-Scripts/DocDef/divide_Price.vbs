Dim price, quantity, lamount, new_lamount

quantity = me.Field("items_ordered").Text
lamount = Replace(me.Field("line_amount").Text,".","")
new_lamount = Replace(lamount,",",".") '63.50
price = new_lamount / quantity

if me.Field("items_ordered").Text <> "" and me.Field("line_amount").Text <> "" then
    me.Field("item_price").Text = price
end if