# Tutorial Install PHPWord di Codeigniter 

hal yang akan kita bahas disini merupakan cara menggunakan PHPWord 0.6.2-1 Beta untuk framework codeigniter.
yang perlu kita perhatikan disini adalah versi PHPWord memiliki struktur yang berbeda disetiap versinya

Namun pada tutorial ini saya menggunkan versi beta (PHPWord 0.6.2-1)

# Apa yang perlu diperhatikan

1. THIRD_PARTY  = Masukkan PHPWord pada folde berikut
2. LIBRARY      = Buat sebuah file baru disini kemudian masukkan kode berikut:
<div class="highlight highlight-source-json"><pre>
<?php if ( ! defined('BASEPATH')) exit('No direct script access allowed');

require_once APPPATH."/third_party/PHPWord.php"; 

class Word extends PHPWord { 
    public function __construct() { 
        parent::__construct(); 
    } 
}
?>
</pre></div>
3. CONTROLLER   = Buat 1 file baru di folder controller kemudian masukkan kode berikut

Example:
<div class="highlight highlight-source-json"><pre>
<?php if (!defined('BASEPATH')) exit ('No direct script access allowed');
class PHPWord_Text extends CI_Controller {
    function __construct() {
        parent::__construct();
        $this->load->library('word');
    }
    function index() {
		
        $PHPWord = $this->word; // New Word Document
        $section = $PHPWord->createSection(); // New portrait section
        // Add text elements
        $section->addText('Hello World!');
        $section->addTextBreak(2);
        $section->addText('Mohammad Rifqi Sucahyo.', array('name'=>'Verdana', 'color'=>'006699'));
        $section->addTextBreak(2);
        $PHPWord->addFontStyle('rStyle', array('bold'=>true, 'italic'=>true, 'size'=>16));
        $PHPWord->addParagraphStyle('pStyle', array('align'=>'center', 'spaceAfter'=>100));
        // Save File / Download (Download dialog, prompt user to save or simply open it)
		$section->addText('Ini Adalah Demo PHPWord untuk CI', 'rStyle', 'pStyle');
		
        $filename='just_some_random_name.docx'; //save our document as this file name
        header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document'); //mime type
        header('Content-Disposition: attachment;filename="'.$filename.'"'); //tell browser what's the file name
        header('Cache-Control: max-age=0'); //no cache
        $objWriter = PHPWord_IOFactory::createWriter($PHPWord, 'Word2007');
        $objWriter->save('php://output');
    }
}
?>

</pre></div>
