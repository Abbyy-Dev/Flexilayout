Dim netamount, vatamount, totalamount, new_total

netamount = Replace(me.Field("invoicenetsum").Text,".","")
totalamount = Replace(me.Field("invoicesum").Text,".","")
new_total = Replace(totalamount,",",".")
new_netamount = Replace(netamount,",",".")
vatamount = new_total - new_netamount

if me.Field("vatamount").Text = "" then
    me.Field("vatamount").Text = vatamount
end if