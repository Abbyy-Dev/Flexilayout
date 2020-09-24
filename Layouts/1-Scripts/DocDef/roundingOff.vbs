================ LATEST!!! 10.00% format ==============

Dim discount, items, price, per, tdiscount, prod, nprod

discount = me.Field("allowancecharge").Text
items = Replace(me.Field("items_ordered").Text,",",".")
price = Replace(me.Field("item_price").Text,",",".")
per = Replace(me.Field("per").Text,",",".")

tdiscount = discount/100
prod = (items*price*per)*(-tdiscount)
nprod = Round(prod,2)

if me.Field("allowancecharge").Text <> "" then

  me.Field("allowancecharge").Text = nprod

end if