Dim email

email = "info@natuuriseenfeest.nl"

if me.Field("supplieremail").Text = "" then

    me.Field("supplieremail").Text = email
    
end if


=============== changing en to / paymentreference =================

Dim payref

payref = Replace(me.Field("paymentreference").Text," en ","/")


if me.Field("paymentreference").Text <> "" then 
    me.Field("paymentreference").Text = payref
end if