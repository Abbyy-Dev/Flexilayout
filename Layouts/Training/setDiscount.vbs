Dim dis, tdis, io, ip, tip, ntip, prod

dis = me.Field("allowancecharge").Text
io = me.Field("items_ordered").Text
ip = me.Field("item_price").Text

tdis = Replace(dis,",",".")/100
tip = Replace(ip,",",".")
ntip = Replace(tip," ","")
prod = (io*ntip)*(-tdis)

if me.Field("allowancecharge").Text <> "" then
  me.Field("allowancecharge").Text = prod
end if