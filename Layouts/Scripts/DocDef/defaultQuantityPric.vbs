================= default items_ordered ==========================

Dim dQuantity

dQuantity = 1

if me.Field("items_ordered").Text = "" then

    me.Field("items_ordered").Text = dQuantity
    
end if


================= default item_price ==========================

Dim dPrice

dPrice = me.Field("line_amount").Text

if me.Field("item_price").Text = "" then

    me.Field("item_price").Text = dPrice
    
end if