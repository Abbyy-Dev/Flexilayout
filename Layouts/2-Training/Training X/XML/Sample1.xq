xquery version "3.1";

declare default element namespace "http://www.globiq.com/profile";

for $masterfile in doc("masterfile.xml")//Employee
let $a := $masterfile/@status
return $masterfile