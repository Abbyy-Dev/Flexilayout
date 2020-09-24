xquery version "3.1";

declare namespace cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2";
declare namespace cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2";
declare default element namespace "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2";

declare function local:totalAmount(
	$lineExtension as xs:decimal?,
	$payableAmount as xs:decimal?) as xs:decimal?
{
	let $sum := $lineExtension + $payableAmount
	return $sum
};
(:
declare function local:num-descendant-elements($el as element()) as xs:decimal?
{
    sum(for $child in $el/*
    return local:num-descendant-elements($child) + 1)
};:)

<try1>
{


let $legalMon := doc("UBL.xml")/Invoice/cac:LegalMonetaryTotal
return local:totalAmount($legalMon/cbc:LineExtensionAmount, $legalMon/cbc:PayableAmount)
}
</try1>
(:
<try2>
{
let $lines := doc("UBL.xml")/Invoice/cac:InvoiceLine
return local:num-descendant-elements($lines/child::*)
}
</try2>:)