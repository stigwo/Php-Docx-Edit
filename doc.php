<!-- 
https://github.com/stigwo

Simple example PHP word .docx online fill

 -->

<?php

$namea = "Seller";
$nameb = "Buyer";
$date = "Date";


?>
 <form class="form" action="" method="post">
                <input type="text" class="input me-2 bg-white" name="namea" placeholder="Seller" value="<?php echo $namea; ?>" required />
				 <input type="text" class="input me-2 bg-white" name="nameb" placeholder="Buyer" value="<?php echo $nameb; ?>" required />
				  <input type="text" class="input me-2 bg-white" name="date" placeholder="Date" value="<?php echo $date; ?>" required />
                <button type="submit" name="search" class="btn bg-dark">
                   <i class="fa fa-search ms-1"></i>
                </button>
              </form>
<?php
if (isset($_POST['search'])) {
    


 $namea = $_POST['namea'];
 $nameb = $_POST['nameb'];
 $date = $_POST['date'];

$input        = 'non.docx';
$output       = 'non-filled.docx';
$replacements = [
    'NAVNKINA' => $namea,
    'NAVNNORGE' => $nameb,
    'DATESIGN' => $date,
    'MANUFCONTACT' => 'Elon Musk',
	'BUYERCONTACT' => 'Donald Trump',
];

$successful = searchReplaceWordDocument($input, $output, $replacements);
 
echo $successful ? "New filled file created $output" : 'Failed!';
}
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