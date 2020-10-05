xquery version "1.0";

declare function local:prod2ndDigit($prod as element()?) as xs:string? {
    substring($prod/number, 2, 1)
};
doc("catalog.xml")//product[local:prod2ndDigit(.) > '5']