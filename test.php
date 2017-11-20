<?php
$data=array(1,2,3,4,5);
foreach($data as $key => $value){
    echo "[$key] is $value</br>";
  
}
$foo1="bar1";
$foo2="bar2";
$foo3="bar3";
$foo4="bar4";

echo "foo1 is $foo1<br/>foo2 is $foo2<br/>foo3 is $foo3<br/>foo4 is $foo4<br/>";


$data=array(1,2,3,4,5,6);
function plus2($x){
    return $x+2;
}
function li($x){
    return "<li>$x</li>";
}
?>

// not bad
<ul>
<?php 
    foreach($data as $key=>$value){
        echo '<li>'.plus2($value).'</li>';
    }
?>
</ul>

// more clear
<ul>
<?php 
    $plus2Data=array_map("plus2",$data);
    $wrapLi=array_map("li",$plus2Data);
    $result=join("",$wrapLi);
    echo $result;
?>
</ul>