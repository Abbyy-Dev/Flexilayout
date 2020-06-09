Dim payref,num

payref = me.Field("paymentreference").Text
num = me.Field("invoicenum").Text

if me.Field("invoicenum").Text <> "" then 
    me.Field("paymentreference").Text = num + "/" + payref
end if