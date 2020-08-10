Dim price, p, per, new_price

price = me.Field("item_price").Text
p = Replace(me.Field("item_price").Text,",",".")
per = me.Field("per").Text
new_price = p / per

if me.Field("per").Text <> "" then
    me.Field("item_price").Text = new_price
else
    me.Field("item_price").Text = price  
end if