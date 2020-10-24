Dim quan, per, new_quan, price, charge

quan = me.Field("items_ordered").Text
per = me.Field("per").Text
price = Replace(me.Field("item_price").Text,".",".")
charge = Replace(me.Field("allowancecharge").Text,".","")

discount = Replace(charge,",",".")/100

new_disc = quan / per * price * - discount


if me.Field("allowancecharge").Text <> "" then
    me.Field("allowancecharge").Text = new_disc
end if