<?php
    $data=array(
        "first"=>array("Volvo",22,18),
        "sec"=>array("BMW",15,13),
        "third"=>array("Saab",5,2),
        "forth"=>array("Land Rover",17,15)
    );
?>


<table border="2" width='200'>            
<tr>
    <th>Name</th>
    <th>Stock</th>
    <th>Sold</th>
</tr>
<?php

$keys=array_keys($data);
for($i=0;$i<count($keys);$i++){
    $key=$keys[$i];
    $row=$data[$key];
    echo "<tr>";
    for ($j=0; $j<count($row); $j++) {
        $col=$row[$j];
        echo "<td>$col</td>";
    }
    echo "</tr>";
}

?>       
</table>

<br>

<table border="2" width='200'>
    <tr>
        <th>Name</th>
        <th>Stock</th>
        <th>Sold</th>
    </tr>
    <?php
    foreach($data as $row){
        echo "<tr>";
        foreach($row as $col){
            echo "<td>$col</td>";
        }
        echo "<tr>";            
    }
    ?>
</table>

<br>

<?php
function parseRow($row)
{
    $row=join(array_map("parseCol", $row));
    return "<tr>$row</tr>";
}
function parseCol($col)
{
    return "<td>$col</td>";
}
$table=join("", array_map("parseRow", $data));
?>

<table border="2" width='200'>
    <tr>
        <th>Name</th>
        <th>Stock</th>
        <th>Sold</th>
    </tr>
    <?php echo $table; ?> 
</table>

