Dim dis, tdis, io, ip, tip, prod, nprod

dis = me.Field("allowancecharge").Text
io = me.Field("items_ordered").Text
ip = me.Field("item_price").Text

tdis = Replace(dis,",",".")/100
tip = Replace(ip,",",".")
prod = (io*tip)*(-tdis)
nprod = Round(prod,2)

if me.Field("allowancecharge").Text <> "" then
  me.Field("allowancecharge").Text = nprod
end if