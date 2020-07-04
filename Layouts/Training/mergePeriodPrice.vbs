===============  4,3452 Weken  =================

Dim price, period, nprice, nperiod, totalPrice


price = Replace(me.Field("item_price").Text,".","")
period = Replace(me.Field("period").Text,".","")
nprice = Replace(price,",",".")
nperiod = Replace(period,",",".")

totalPrice = nprice * nperiod

rprice = Round(totalPrice,5)

if me.Field("item_price").Text <> "" then

  me.Field("item_price").Text = rprice

end if

