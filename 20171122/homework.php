<?php
if(empty($_POST["height"])){
    echo "please type in all information"."<br>";
}
else if(empty($_POST["weight"])){
    echo "please type in all information"."<br>";
}
else{
echo "height =".$_POST["height"]."<br>";
echo "weight =".$_POST["weight"]."<br>";
echo "BMI:".($_POST["weight"]/(($_POST["height"]/100)*($_POST["height"]/100)))."<br>";}

if($_FILES["file"]["error"]==4){
    echo "empty"."<br>"; 
}
else if ((($_FILES["file"]["type"] == "image/gif")
|| ($_FILES["file"]["type"] == "image/jpeg")
|| ($_FILES["file"]["type"] == "image/jpg")
|| ($_FILES["file"]["type"] == "image/x-png")
|| ($_FILES["file"]["type"] == "image/png"))){
    $filename = $_FILES["file"]["name"];
    move_uploaded_file($_FILES["file"]["tmp_name"],"upload/".$filename);
    echo '<img src="upload/'.$filename.'"/>';
    }
else{
    echo "wrong file type"."<br>";
}
?>