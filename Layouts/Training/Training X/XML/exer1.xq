xquery version "3.1";

declare default element namespace "http://www.globiq.com/profile";

<Root>
<item1>{
for $emp in doc("Masterfile.xml")//Employee
where $emp/@*[.="Regular"]
return 
if ( $emp/@*[.="Regular"] and $emp/Address != "") then
<result>{
$emp/Name, $emp/Address, $emp/ContactNumber
}</result>
else "Employee must have a current/permanent address."
}</item1> |


<item2>{
for $emp in doc("Masterfile.xml")//Employee
let $first := substring-after($emp/Name,", "),
	$last := substring-before($emp/Name,", ")
order by $emp/@Status
return
if ( $emp/Name !="" ) then
<result>
<Firstname>{$first}</Firstname>
<Lastname>{$last}</Lastname>
{$emp/child::*[position()>2]}

</result>

else "Employee must have a name."
}</item2> |

<item3>{
for $emp in doc("Masterfile.xml")//Employee
let $email := $emp/EmailAddress,	
	$personal := concat("+63",$emp/ContactNumber/Personal),
	$work := concat("+63",$emp/ContactNumber/Work),
	$child := count($emp/ContactNumber/child::*)
where $emp/ContactNumber != ""
return 
if ( $child = 0 ) then
<result>{$email,<ContactNumber>{concat("+63",$emp/ContactNumber)}</ContactNumber>}</result>
else if ( $child = 2 ) then
<result>{$email, <ContactNumber>
<Personal>{$personal}</Personal>
<Work>{$work}</Work></ContactNumber>}</result>
else "No Contact Number"
}</item3>

<item4>{
for $emp in doc("Masterfile.xml")//Employee
let $age := $emp/Birthday,
	$year := substring($age,7,4),
	$month := substring($age,4,2),
	$day := substring($age,1,2)
return
<result>{
$emp/Name,
<Birthday>{concat($year,"-",$month,"-",$day)}</Birthday>
}</result>
}</item4>

<item5>{
for $emp in doc("Masterfile.xml")//Employee
let $age := $emp/Birthday,
	$year := substring($age,7,4),
	$month := substring($age,4,2),
	$day := substring($age,1,2),
	$bday := xs:date(concat($year,"-",$month,"-",$day))
return
<result>{$emp/Name,
<Age>{altovaext:age($bday)}</Age>
}</result>

}</item5>

</Root>