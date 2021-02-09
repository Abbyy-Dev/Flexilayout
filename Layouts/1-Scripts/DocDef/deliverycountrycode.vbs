if me.Field("deliverycountrycode").Text = "Nederland" or me.Field("deliverycountrycode").Text = "Nederland" then
me.Field("deliverycountrycode").Text = "NL"
end if


================== Update 2021-01-30 ======================
if me.Field("deliverycity").Text <> "" then
me.Field("deliverycountrycode").Text = "NL"
end if