<?php
error_reporting(E_ALL);
ini_set('display_errors', TRUE); 
ini_set('display_startup_errors', TRUE); 
date_default_timezone_set('Europe/London');

/** Include PHPExcel */
require_once dirname(__FILE__) . '/Classes/PHPExcel.php';



function colorInverse($color){
    $color = str_replace('#', '', $color);
    if (strlen($color) != 6){ return '000000'; }
    $rgb = '';
    for ($x=0;$x<3;$x++){
        $c = 255 - hexdec(substr($color,(2*$x),2));
        $c = ($c < 0) ? 0 : dechex($c);
        $rgb .= (strlen($c) < 2) ? '0'.$c : $c;
    }
    return $rgb;
}


function lockCells(&$lockedCells, $MergedCells)
{
    
    $myArray = explode(':', $MergedCells);
    
    $BeginCell = $myArray[0];
    $EndCell   = $myArray[1];
    
    $BeginCol="";
    $BeginRow="";
    
    for ($i = 0; $i < strlen($BeginCell); $i++)
    {
        $asciiValue = ord(strtolower($BeginCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $BeginCol = $BeginCol.$BeginCell{$i};
        }else{
                 $BeginRow = $BeginRow.$BeginCell{$i};
             };
    }
    
    $BeginColNumber = PHPExcel_Cell::columnIndexFromString($BeginCol);
    
    $EndCol="";
    $EndRow="";
    
    for ($i = 0; $i < strlen($EndCell); $i++)
    {
        $asciiValue = ord(strtolower($EndCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $EndCol = $EndCol.$EndCell{$i};
        }else{
                 $EndRow = $EndRow.$EndCell{$i};
             };
    }

    $EndColNumber = PHPExcel_Cell::columnIndexFromString($EndCol);
              
    if ( ( $BeginColNumber != $EndColNumber ) && ( $BeginRow == $EndRow ) )    
    { 
        $MergeType="horizontal"; 
        for ( $ColName=$BeginCol; $ColName <= $EndCol; $ColName++)
        {
            array_push($lockedCells, "$ColName$BeginRow");
        }
        #echo "ZABLOKOWANE KOMORKI (horizontal) : \n";
        #print_r($lockedCells);
    }
    else if ( ( $BeginColNumber == $EndColNumber ) && ( $BeginRow != $EndRow ) ) 
    {
        $MergeType="vertical";
        for ( $ColName=$BeginRow; $ColName <= $EndRow; $ColName++)
        {
            array_push($lockedCells, "$BeginCol$ColName");
        }
       # echo "ZABLOKOWANE KOMORKI : (vertical) \n";
       # print_r($lockedCells);
       # echo "<br>";

    }

    else if ( ( $BeginColNumber != $EndColNumber ) && ( $BeginRow != $EndRow ) )    { $MergeType="both"; }        
    else                                                                { $MergeType="WTF ??? Merge on one cell ???"; };
 
}

function isLocked(&$lockedCells, $Cell)
{
    $status=false;
    if (in_array($Cell, $lockedCells)) 
    {
        $status=true;
    }else{
             $status=false;
         };

    return $status;
}

function GetMergedCellsCount($MergedCells, $MergeType)
{
    $myArray = explode(':', $MergedCells);
    
    $BeginCell = $myArray[0];
    $EndCell   = $myArray[1];
    $BeginCol="";
    $BeginRow="";
    
    for ($i = 0; $i < strlen($BeginCell); $i++)
    {
        $asciiValue = ord(strtolower($BeginCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $BeginCol = $BeginCol.$BeginCell{$i};
        }else{
                 $BeginRow = $BeginRow.$BeginCell{$i};
             };
    }
    
    $BeginColNumber = PHPExcel_Cell::columnIndexFromString($BeginCol);

    $EndCol="";
    $EndRow="";
    
    for ($i = 0; $i < strlen($EndCell); $i++)
    {
        $asciiValue = ord(strtolower($EndCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $EndCol = $EndCol.$EndCell{$i};
        }else{
                 $EndRow = $EndRow.$EndCell{$i};
             };
    }

    $EndColNumber = PHPExcel_Cell::columnIndexFromString($EndCol);
    
    if ( $MergeType == "horizontal" )    return ( $EndColNumber - $BeginColNumber + 1 );
    if ( $MergeType == "vertical" )      return ( $EndRow - $BeginRow + 1 );
    if ( $MergeType == "both" )          return "Ehm... I don't know if redmine accepting this type of merge... :(";
}

function GetMergeType($MergedCells)
{
    $myArray = explode(':', $MergedCells);
    
    $BeginCell = $myArray[0];
    $EndCell   = $myArray[1];
    
    #echo "Begin Cell : $BeginCell\n  End Cell : $EndCell\n";
    
    $BeginCol="";
    $BeginRow="";
    
    for ($i = 0; $i < strlen($BeginCell); $i++)
    {
        $asciiValue = ord(strtolower($BeginCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $BeginCol = $BeginCol.$BeginCell{$i};
        }else{
                 $BeginRow = $BeginRow.$BeginCell{$i};
             };
    }
    
    $BeginColNumber = PHPExcel_Cell::columnIndexFromString($BeginCol);
    
    #echo "BC : ".$BeginColNumber." (".$BeginCol.")\nBR : ".$BeginRow."\n";

    $EndCol="";
    $EndRow="";
    
    for ($i = 0; $i < strlen($EndCell); $i++)
    {
        $asciiValue = ord(strtolower($EndCell{$i}));
        if ($asciiValue > 96 && $asciiValue < 123)
        {
            $EndCol = $EndCol.$EndCell{$i};
        }else{
                 $EndRow = $EndRow.$EndCell{$i};
             };
    }

    $EndColNumber = PHPExcel_Cell::columnIndexFromString($EndCol);
    
   
    #echo "EC : ".$EndColNumber." (".$EndCol.")\nER : ".$EndRow."\n";
              
    if      ( ( $BeginColNumber != $EndColNumber ) && ( $BeginRow == $EndRow ) )    { $MergeType="horizontal"; }
    else if ( ( $BeginColNumber == $EndColNumber ) && ( $BeginRow != $EndRow ) )    { $MergeType="vertical"; }
    else if ( ( $BeginColNumber != $EndColNumber ) && ( $BeginRow != $EndRow ) )    { $MergeType="both"; }        
    else                                                                { $MergeType="WTF ??? Merge on one cell ???"; };
    
    return $MergeType;
}

?>
