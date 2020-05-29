Dim dis, tdis, io, ip, tip, prod

dis = me.Field("allowancecharge").Text
io = me.Field("items_ordered").Text
ip = me.Field("item_price").Text

tdis = Replace(dis,",",".")/100
tip = Replace(ip,",",".")
prod = (io*tip)*(-tdis)

if me.Field("allowancecharge").Text <> "" then
  me.Field("allowancecharge").Text = prod
end if
