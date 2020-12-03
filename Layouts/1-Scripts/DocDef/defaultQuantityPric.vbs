================= default items_ordered ==========================

Dim dQuantity

dQuantity = 1

if me.Field("items_ordered").Text = "" then

    me.Field("items_ordered").Text = dQuantity
    
end if


================= default item_price ==========================

Dim dPrice

dPrice = me.Field("line_amount").Text

if me.Field("item_price").Text = "" then

    me.Field("item_price").Text = dPrice
    
end if

================= default vatbase ==========================

Dim dVatbase

dVatbase = me.Field("invoicenetsum").Text

if me.Field("vatbase").Text = "" then

    me.Field("vatbase").Text = dVatbase
    
end if

================= default set deliverycountrycode ==========================

if me.Field("deliverycity").Text <> "" then

    me.Field("deliverycountrycode").Text = "NL"
    
end if

================= default line_amount ==========================

Dim price, tprice, quantity, tquantity, lineamount

tprice = Replace(me.Field("item_price").Text,".","") '2500,00
price = Replace(tprice,",",".") '2500.00

tquantity = Replace(me.Field("items_ordered").Text,".","")
quantity = Replace(tquantity,",",".")

lineamount = Round(quantity * price,3)

if me.Field("line_amount").Text = "" then
    me.Field("line_amount").Text = lineamount
end if

================= default invoicenetsum ==========================

Dim dNet, tNet

vat = Replace(me.Field("vatamount").Text,",",".")
total  = Replace(me.Field("invoicesum").Text,",",".")
tNet = total - vat
dNet = Round(tNet,5)

if me.Field("invoicenetsum").Text = "" then

    me.Field("invoicenetsum").Text = dNet
    
end if


================= set allowancecharge if amount is zero ==========================

Dim reason

reason = me.Field("reason").Text


if me.Field("amount").Text = "0,00" then

    me.Field("reason").Text = ""
 
else
 me.Field("reason").Text = reason
    
end if