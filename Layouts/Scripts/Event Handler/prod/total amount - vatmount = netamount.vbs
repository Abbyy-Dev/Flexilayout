Dim netamount, vatamount, totalamount, new_total, new_vatamount

vatamount = Replace(me.Field("vatamount").Text,".","")
totalamount = Replace(me.Field("invoicesum").Text,".","")
new_total = Replace(totalamount,",",".")
new_vatamount = Replace(vatamount,",",".")
netamount = new_total - new_vatamount

if me.Field("invoicenetsum").Text = "" then
    me.Field("invoicenetsum").Text = netamount
end if

