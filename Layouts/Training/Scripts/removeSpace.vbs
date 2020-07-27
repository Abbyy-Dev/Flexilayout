Dim curr_date , tr_date

curr_date = CStr(me.Field("orderdatebuyer").Text)
tr_date = Trim(curr_date)


if me.Field("orderdatebuyer").Text <> "" then
   me.Field("orderdatebuyer").Text = tr_date
end if