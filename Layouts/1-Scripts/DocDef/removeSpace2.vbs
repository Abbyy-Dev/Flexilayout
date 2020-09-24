Dim ctotal, ntotal

ctotal = me.Field("totalorderamount").Text
ntotal = Replace(ctotal," ","")

if me.Field("totalorderamount").Text <> "" then
me.Field("totalorderamount").Text = ntotal
end if