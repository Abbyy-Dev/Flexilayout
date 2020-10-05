xquery version "3.1";

declare namespace cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
declare namespace cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
declare default element namespace "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2";



declare function local:totalAmount(
	$lineExtension as xs:decimal?,
	$payableAmount as xs:decimal?) as xs:decimal?
{
	let $sum1 := $lineExtension + $payableAmount
	return $sum1
};


declare function local:total(
	$lineExt as xs:decimal?,
	$totalTax as xs:decimal?,
	$payable as xs:decimal?) as xs:decimal?
{
	let $sum2 := $lineExt + $totalTax
		
	return
	
	if ( $sum2 = $payable ) then
		$sum2
	else 000

};


declare function local:sumLegal(
	$child as element(cac:LegalMonetaryTotal)) as xs:double
	
{
	let $sum3 := sum($child/child::*)
	return $sum3
};




<try1>
{
let $legalMon := doc("UBL.xml")/Invoice/cac:LegalMonetaryTotal
return local:totalAmount($legalMon/cbc:LineExtensionAmount, $legalMon/cbc:PayableAmount)
}
</try1>|

<try2>
{
let $invoice := doc("UBL.xml")/Invoice

return local:total($invoice/cac:LegalMonetaryTotal/cbc:LineExtensionAmount, $invoice/cac:TaxTotal/cbc:TaxAmount, $invoice/cac:LegalMonetaryTotal/cbc:PayableAmount)
}
</try2>|

<try3>
{
let $sumLegal := doc("UBL.xml")/Invoice
return local:sumLegal($sumLegal/cac:LegalMonetaryTotal)
}
</try3>