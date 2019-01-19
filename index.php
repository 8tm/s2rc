<?php
    echo '<?xml version="1.0" encoding="iso-8859-2"?>';
    $DEBUG=isset($_GET['debug']);
    if ($DEBUG)
    {
        echo "DEBUG MODE ON<br>";
    }
?>
<!DOCTYPE html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="icon" href="template/images/favicon.ico">
    <title>sheet2redmine converter</title>
    <!-- Bootstrap core CSS -->
    <link href="template/css/bootstrap.min.css" rel="stylesheet">

    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <link href="template/css/ie10-viewport-bug-workaround.css" rel="stylesheet">

    <!-- Custom styles for this template -->
    <link href="template/css/navbar-fixed-top.css" rel="stylesheet">

    <!-- Just for debugging purposes. Don't actually copy these 2 lines! -->
    <!--[if lt IE 9]><script src="ie8-responsive-file-warning.js"></script><![endif]-->
    <script src="template/js/ie-emulation-modes-warning.js"></script>

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.3/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>

  <body>

    <!-- Fixed navbar -->
    <nav class="navbar navbar-default navbar-fixed-top">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
            <span class="sr-only">navigation</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <a class="navbar-brand" href="http://www.s2rc.tmi.ovh">s2rc</a>
          <a href=""></a></li>
        </div>
        <div id="navbar" class="navbar-collapse collapse">
          <ul class="nav navbar-nav">
          </ul>
          <ul class="nav navbar-nav navbar-right">
            <!--<li><a href="../navbar/">Default</a></li>-->
            <!--<li><a href="../navbar-static-top/">Static top</a></li>-->
            <li class="active"><a href="./">Wróć<span class="sr-only">(current)</span></a></li>
          </ul>
        </div><!--/.nav-collapse -->
      </div>
    </nav>

    <div class="container">

      <!-- Main component for a primary marketing message or call to action -->
      <div class="jumbotron">

        <h2>sheet2redmine converter</h2>

        <div>
            <form enctype="multipart/form-data" action="./<?php if ($DEBUG) { echo "?debug"; } ?>" method="POST">
                <input type="hidden" name="MAX_FILE_SIZE" value="50000" />
                <input class="btn btn-lg btn-primary" name="plik" type="file" />
                <input class="btn btn-lg btn-primary" type="submit" value="Wyślij plik" />
            </form>
        </div>

<p>LibreOffice Calc :</p>
<div>
* Uruchamiamy LibreOffice Calc i tworzymy własną tabelę<br>
* Zapisujemy do Excel 2003 (*.xls) z użyciem formatu MS<br>
* Wgrywamy plik<br>
* Klikamy "Wyślij plik"<br>
* Klikamy w wygenerowany tekst (zaznaczy się cały tekst) i go kopiujemy<br>
* Wklejamy tekst w Redmine<br>
</div><br>
<p>Aplikacja obecnie obsługuje :</p>
<p><div> * obliczanie formuł ( np : "=sum(A7:K17)" )<br> * scalanie komórek<br> * pogrubione<br> * kursywa<br> * kolor czcionki<br> * kolor tła<br>* justowanie tekstu do prawej / do środka / standardowo po lewej</div></p>

           <p>Tabela w LibreOffice Calc zapisana do EXCEL 2003 (*.xls) : </p>
           <p><div>Po lewej : LibreOffice Calc (EXCEL 2003)</div><div>Po prawej : Redmine table: </div></p>
           <img height="264px" src="template/images/xls-screen.png" alt="LibreOffice Calc - PODGLĄD PLIKU XLS">
           <img height="264px" src="template/images/redmine-screen.png" alt="Redmine - wygląd tabelki z pliku XLS"><br>
        <p>Przykładowy plik xls można pobrać tutaj : <a href="examples/test_merged_cells.xls">test_merged_cells.xls</a>
        <p>

        <div>
    <?php
        $Prefix=date('Y-m-d____G-i-s')."___";
        $plik_tmp = $_FILES['plik']['tmp_name'];
        $plik_nazwa = $_FILES['plik']['name'];
        $plik_rozmiar = $_FILES['plik']['size'];

        if(is_uploaded_file($plik_tmp))
        {

            move_uploaded_file($plik_tmp, "ods_xls/$Prefix$plik_nazwa");
            echo "Redmine Table based on file : \"<strong>$plik_nazwa</strong>\"<br>";
            error_reporting(E_ALL);
            ini_set('display_errors', TRUE); 
            ini_set('display_startup_errors', TRUE); 
            date_default_timezone_set('Europe/Warsaw');
            define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');


            require_once dirname(__FILE__) . '/Classes/PHPExcel.php';
            require_once dirname(__FILE__) . '/functions.php';


            $objPHPExcel = new PHPExcel();


            $sFilepath="ods_xls/$Prefix$plik_nazwa";
            $objPHPExcel = PHPExcel_IOFactory::load($sFilepath);

            $sheet = $objPHPExcel->setActiveSheetIndex(0); 

            $dimension = $sheet->calculateWorksheetDimension();  # echo "Rozmiar          : ".$dimension."<br>";
            $highestColumn = $sheet->getHighestColumn();         # echo "Ostatnia kolumna : ".$highestColumn."<br>";
            $highestRow    = $sheet->getHighestRow();            # echo "Ostatni wiersz   : ".$highestRow ."<br>";
            $last_colek="";
            $redmine_str="\n";

            $lockedCells = array();
            for ( $rowek=1; $rowek <= $highestRow; $rowek++)
            {

                for ( $colek='A'; $colek<=$highestColumn; $colek++)
                {


                    $IsLocked=false;
                    $isBcolor=false;
                    $isBcolor=false;
                    $isBold=false;
                    $isItalic=false;
                    $CellCol="$colek";
                    $CellRow="$rowek";
                    $CellAddress="$CellCol$CellRow";
                    $cell = $sheet->getCell($CellAddress);
                    $v = $cell->getCalculatedValue();    # Było : $v = $cell->getValue();
                    $backgroundColor = $sheet->getCell($CellAddress)->getStyle($CellAddress)->getFill()->getStartColor()->getRGB(); 
                    $fontColor = $sheet->getCell($CellAddress)->getStyle($CellAddress)->getFont()->getColor()->getRGB(); 
                    $Bold = $sheet->getCell($CellAddress)->getStyle($CellAddress)->getFont()->getBold(); 
                    $Italic = $sheet->getCell($CellAddress)->getStyle($CellAddress)->getFont()->getItalic(); 
                    $Align = $sheet->getCell($CellAddress)->getStyle($CellAddress)->getAlignment()->getHorizontal();
                    #000000 - Black
                    #FFFFFF - White


                              if ($fontColor != "000000")
                              {
                                  $fColorB = "%{color:#".$fontColor."}";
                                  $fColorE="%";
                              }else{
                                       $fColorB = "";
                                       $fColorE = "";
                                   };
                              if ($backgroundColor != "000000") { $bColor  = "{background:#".$backgroundColor."}";                 }else{ $bColor  = "";         };



                    if ($Bold != "")                  { $bold    = "*";                                                  }else{ $bold    = "";                   };
                    if ($Italic != "")                { $italic  = "_";                                                  }else{ $italic  = "";                   };

                    if ($Align == "center")
                    {
                        $align  = "=";
                    }else if ($Align == "right")
                    {
                        $align  = ">";
                    }else{
                              $align="";
                         };

                    if (! isLocked($lockedCells, $CellAddress) )
                    {
                        foreach ($sheet->getMergeCells() as $cells)
                        {
                            if ($cell->isInRange($cells))
                            {
                                $MergedCells=$cells;
                                lockCells($lockedCells, $cells);
                                $IsLocked=true;
                                if ( GetMergeType($cells) == "horizontal" )
                                {
                                    $redmine_str=$redmine_str."|\\".GetMergedCellsCount( $cells, GetMergeType($cells) ).$align.$bColor.". ".$italic.$bold.$fColorB.$v.$fColorE.$bold.$italic;

                                }else if ( GetMergeType($cells) == "vertical" )
                                      {

                                          $redmine_str=$redmine_str."|/".GetMergedCellsCount( $cells, GetMergeType($cells) ).$align.$bColor.". ".$italic.$bold.$fColorB.$v.$fColorE.$bold.$italic;

                                      }else{ 
                                               echo "Nieobsługiwane w redmine złączenie komórek (poziome i pionowe) w : $cells";
                                               break; 
                                           };
                                break;
                            }else{
                                     $IsLocked=false;
                                 };
                        }

                        if ( ! $IsLocked )
                        {
                            #$redmine_str=$redmine_str."| $v ";
                            $redmine_str=$redmine_str."|".$align.$bColor.". ".$italic.$bold.$fColorB.$v.$fColorE.$bold.$italic;
                        }
                    }

                    $last_colek=$colek;

                }
                $redmine_str=$redmine_str."|\n";
            }
            echo "<textarea class=\"js-copytextarea\" name=\"redmine_table\" rows=\"$highestRow\" onclick=\"this.focus();this.select()\" readonly=\"readonly\">".$redmine_str."</textarea>";

}
?>
    </div>
        </p>
      </div>
    </div> <!-- /container -->


    <!-- Bootstrap core JavaScript
    ================================================== -->
    <!-- Placed at the end of the document so the pages load faster -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.12.4/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="template/js/jqueryx.min.js"><\/script>')</script>
    <script src="template/js/bootstrap.min.js"></script>
    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="template/js/ie10-viewport-bug-workaround.js"></script>
  </body>
</html>

