Dim linenet, rlinenet, kwcredit, newline

linenet = me.Field("line_amount").Text
rlinenet = Replace(me.Field("line_amount").Text,"CR","-") 
newline = Replace(rlinenet,"-","")
kwcredit = Right(rlinenet,1)

if me.Field("line_amount").Text <> "" and kwcredit = "-" then
    me.Field("line_amount").Text = "-" & newline
else
    me.Field("line_amount").Text = linenet
end if


============ move negative to the left=====



Dim quantity, nquantity, neg

quantity = me.Field("items_ordered").Text
nquantity = Replace(quantity,"-","")
neg = Right(quantity,1)

if me.Field("items_ordered").Text <> "" and neg = "-" then
    me.Field("items_ordered").Text = "-" + nquantity
else
    me.Field("items_ordered").Text = quantity
end if
    