Dim dQuantity,cQuantity

dQuantity = Replace(me.Field("items_ordered").Text,",",".")
cQuantity = CDbl(dQuantity) / 100

if me.Field("items_ordered").Text <> "" then

    me.Field("items_ordered").Text = cQuantity
    
end if