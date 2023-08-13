<!-- 
https://github.com/stigwo

Simple example PHP word .docx online fill

 -->

<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <!-- Favicon -->
  <link rel="shortcut icon" href="images/favicon.png" type="image/x-icon">
  <!-- FontAwesome -->
  <link rel="stylesheet" href="css/all.min.css" />
  <!-- Bootstrap CSS -->
  <link rel="stylesheet" href="css/bootstrap.min.css" />
  <!-- Custom CSS -->
  <link rel="stylesheet" href="css/styleip.css" />
  <title>Fill it !</title>
</head>

<body class="bg-default">
        
    <!-- Header Start -->
        <header>
          <!-- Navbar -->
          <nav class="navbar navbar-expand lg navbar-light bg-dark shadow-sm">
            <div class="container-fluid">
              <div class="collapse navbar-collapse">
                <ul class="navbar-nav m-auto mb-2 mb-lg-0">
                  <li class="nav-item active">
                    <a class="nav-link text-white" aria-current="page" href="#home"><i class="fa fa-home"></i>&nbsp;Home</a>
                  </li>
                  <li class="nav-item">
                    <a class="nav-link text-white" href="#howtouse"><i class="fa fa-user">&nbsp;</i>How To Use</a>
                  </li>
                </ul>
              </div>
            </div>
          </nav>
          <!-- Navbar -->
        </header>
    <!-- Header End -->


<?php


$date = date("F.d.Y");


?>
<div class="container">
    <!-- Search Start -->
      <div class="search">
        <div class="container">
          <div class="row">
            <div class="col-md-5 mx-auto">
              <div class="txt text-center">
               
                   <img src="images/favicon.png" class="w-25 logo" />
                    <h1 class="main-heading">Fill it !</h1>
               
                <p class="sub-heading">Fill contracts and NDA's</p>
              </div>
 <form  action="" method="post">
  <div class="form-group">
  <label for="exampleFormControlInput1">Giver company of the NDA</label>
                <input type="text" class="form-control" name="namea" placeholder="Company" value="" required />
				</div>
				<div class="form-group">
  <label for="exampleFormControlInput1">Giver person of the NDA</label>
                <input type="text" class="form-control" name="persona" placeholder="Name" value="" required />
				</div>
				 <div class="form-group">
				 <label for="exampleFormControlInput1">Reciever company of the NDA</label>
				 <input type="text" class="form-control" name="nameb" placeholder="Company" value="" required />
				 </div>
				 		<div class="form-group">
  <label for="exampleFormControlInput1">Reciver person of the NDA</label>
                <input type="text" class="form-control" name="personb" placeholder="Name" value="" required />
				</div>
				  <div class="form-group">
				  <label for="exampleFormControlInput1">Date</label>
				  <input type="text" class="form-control" name="date" placeholder="Select Date" value="<?php echo $date; ?>" required />
				  </div>
				  <br>
				   <div class="form-group">
                <button type="submit" name="search" class="btn btn-dark">Fill it !</button>
				</div>
			 </form>
			              </div>
          </div>
      </div>
	  
<?php
if (isset($_POST['search'])) {
    


 $namea = $_POST['namea'];
 $nameb = $_POST['nameb'];
 $persona = $_POST['persona'];
 $personb = $_POST['personb'];
 $date = $_POST['date'];

$input        = 'org/non.docx';
function random_string($length) {
    $key = '';
    $keys = array_merge(range(0, 9), range('a', 'z'));

    for ($i = 0; $i < $length; $i++) {
        $key .= $keys[array_rand($keys)];
    }

    return $key;
}
$ranstring = random_string(20);
$ending = ".docx";
$output = 'gen/' . $ranstring . $ending;

//$output       = 'non-filled.docx';
$replacements = [
    'NAVNKINA' => $namea,
    'NAVNNORGE' => $nameb,
    'DATESIGN' => $date,
    'MANUFCONTACT' => $persona,
	'BUYERCONTACT' => $personb,
];

$successful = searchReplaceWordDocument($input, $output, $replacements);
?>
<div class="txt text-center"><p>
<?php 
echo $successful ? "New filled file created $output" : 'Failed!';
//$link = 'http://localhost/doc/';
$link1 = "https://$_SERVER[HTTP_HOST]";
$link2 = "/doc/";
$link = $link1 . $link2;
$outputx = "0";
$outputx = $link . $output;
echo '<a class="btn btn-dark" href="'.$outputx.'" role="button">DOWNLOAD ';
}
?>
<?php  ?>
</a>
</p></div>
<?php
/**
 * Edit a Word 2007 and newer .docx file.
 * Utilizes the zip extension http://php.net/manual/en/book.zip.php
 * to access the document.xml file that holds the markup language for
 * contents and formatting of a Word document.
 *
 * In this example we're replacing some token strings.  Using
 * the Office Open XML standard ( https://en.wikipedia.org/wiki/Office_Open_XML )
 * you can add, modify, or remove content or structure of the document.
 *
 * @param string $input
 * @param string $output
 * @param array  $replacements
 *
 * @return bool
 */
function searchReplaceWordDocument(string $input, string $output, array $replacements): bool
{
    if (copy($input, $output)) {

        // Create the Object.
        $zip = new ZipArchive();

        // Open the Microsoft Word .docx file as if it were a zip file... because it is.
        if ($zip->open($output, ZipArchive::CREATE) !== true) {
            return false;
        }

        // Fetch the document.xml file from the word subdirectory in the archive.
        $xml = $zip->getFromName('word/document.xml');

        // Replace
        $xml = str_replace(array_keys($replacements), array_values($replacements), $xml);

        // Write back to the document and close the object
        if (false === $zip->addFromString('word/document.xml', $xml)) {
            return false;
        }
        $zip->close();

        return true;
    }

    return false;
}
?>

<!-- About App Section-->
<div class="wrapper bg-light p-5" id="howtouse">
    <div class="container">
        <div class="row">
            <h5>About the app:</h5>
            <p>Fill it ! is a free online tool that is easy to use and lets you fill business templates with your data, and download the result as a *.docx Word file.</p>
            <h5>How to Use?</h5>
            <ul>
                <li>1. Enter the information</li>
                <li>2. Click the "Fill" Button</li>
            </ul>
        </div>
    </div>
</div>

 
<!--Footer-->
    <footer class="page-footer bg-white font-small p-2">
    
      <!-- Copyright -->
      
      <div class="footer-copyright text-center py-3">
        <p>Copyright &copy; <?php echo date("Y"); ?> <b>Fill it</b> - All rights reserved</p>
		<p><a href="https://github.com/stigwo">Get the script for free on my Github</p>
      </div>
      <!-- Copyright -->
    
    </footer>
<!-- Footer -->
   <script src="script/jquery-2.2.4.js"></script>
  <!-- Bootstrap Bundle -->
  <script src="script/bootstrap.bundle.min.js"></script>
</body>

</html>