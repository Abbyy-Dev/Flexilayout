Dim ordernum,pos

ordernum = me.Field("ordernum").Text
pos = me.Field("positie").Text

if me.Field("positie").Text <> "" then 
    me.Field("ordernum").Text = ordernum + "-" + pos
end if