Dim discount, quantity, price, t_discount, t_quantity, t_price, ac

'1.500,25
discount = me.Field("allowancecharge").Text
quantity = Replace(me.Field("items_ordered").Text,".","")
price = Replace(me.Field("item_price").Text,".","")

t_discount = discount/100

'1500.25
t_quantity = Replace(quantity,",",".")
t_price = Replace(price,",",".")
ac = (t_quantity*t_price)*(-t_discount)

if me.Field("allowancecharge").Text <> "" then
  me.Field("allowancecharge").Text = ac
end if