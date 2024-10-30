<?php


    class DocxConversion{
    private $filename;

    public function __construct($filePath) {
        $this->filename = $filePath;
    }

    private function read_docx(){

        $striped_content = '';
        $content = '';

        $zip = zip_open($this->filename);

        if (!$zip || is_numeric($zip)) return false;

        while ($zip_entry = zip_read($zip)) {

            if (zip_entry_open($zip, $zip_entry) == FALSE) continue;

            if (zip_entry_name($zip_entry) != "word/document.xml") continue;

            $content .= zip_entry_read($zip_entry, zip_entry_filesize($zip_entry));

            zip_entry_close($zip_entry);
        }// end while

        zip_close($zip);

		$txt = str_replace('</w:r></w:p></w:tc><w:tc>', " ", $content);
        $txt = str_replace('</w:r></w:p>', "\r\n\r\n", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading1"/>', "<h1>", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading2"/>', "<h2>", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading3"/>', "<h3>", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading4"/>', "<h4>", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading5"/>', "<h5>", $txt);
        $txt = str_replace('<w:pStyle w:val="Heading6"/>', "<h6>", $txt);

		$txt = str_replace('<w:b/>', "<b>", $txt);
        $txt = str_replace('<w:i/>', "<i>", $txt);
        $txt = str_replace('</w:t>', "<wt>", $txt);

		$txt = preg_replace("/<(\w+)>(.*?)<wt>/", "<$1>$2</$1>", $txt);
        $striped_txt = strip_tags($txt, "<br><b><i><h1><h2><h3>");

		$striped_txt = preg_replace("/<(\w+)><(\w+)><(\w+)>/", "<$1>", $striped_txt);
		$striped_txt = preg_replace("/<(\w+)><(\w+)>/", "<$1>", $striped_txt);


        return $striped_txt;
    }
    private function read_txt(){

$handle = fopen($this->filename, "r");
$contents = fread($handle, filesize($this->filename));
fclose($handle);

        //$striped_txt = strip_tags($contents);
$striped_txt = utf8_decode(strip_tags($contents));

		
		return $striped_txt;
	}

    public function convertToText() {

        if(isset($this->filename) && !file_exists($this->filename)) {
            return "File Not exists";
        }

        $fileArray = pathinfo($this->filename);
        $file_ext  = $fileArray['extension'];

         if($file_ext == "docx") {
                return $this->read_docx();
            
        } else if($file_ext == "txt") {
                return $this->read_txt();
            
        } else {
            return "Invalid File Type";
        }
    }

}



if (isset($_POST['create']) && $_FILES['fileToUpload']["tmp_name"][0]) {

$target_dir = wp_upload_dir();

if (!file_exists($target_dir['basedir'] . '/docx2post')) {
    mkdir($target_dir['basedir'] . '/docx2post', 0775, true);
}

function reArrayFiles(&$file_post) {

    $file_ary = array();
    $file_count = count($file_post['name']);
    $file_keys = array_keys($file_post);

    for ($i=0; $i<$file_count; $i++) {
        foreach ($file_keys as $key) {
            $file_ary[$i][$key] = $file_post[$key][$i];
        }
    }

    return $file_ary;
}

if ($_FILES['fileToUpload']) {
    $file_ary = reArrayFiles($_FILES['fileToUpload']);

    foreach ($file_ary as $file) {



$target_file = $target_dir['basedir'] ."/docx2post/". basename($file["name"]);

move_uploaded_file($file["tmp_name"], $target_file);
$filename = $target_dir['basedir'] ."/docx2post/". basename($file["name"]);


    $user_id = get_current_user_id();
	$title = pathinfo( $file["name"], PATHINFO_FILENAME);
	$post_type = $_POST['type'];
	$post_category = explode("||",$_POST['category']);

$docObj = new DocxConversion($filename);
$content= $docObj->convertToText();

    $defaults = array(
        'post_author' => $user_id,
        'post_content' => $content,
        'post_content_filtered' => '',
        'post_title' => $title,
        'post_excerpt' => '',
        'post_status' => 'publish',
        'post_type' => $post_type,
        'comment_status' => '',
        'ping_status' => '',
        'post_password' => '',
        'to_ping' =>  '',
        'pinged' => '',
        'post_parent' => 0,
        'menu_order' => 0,
        'guid' => '',
        'import_id' => 0,
        'context' => '',
    );
 
    $postarr = wp_parse_args($postarr, $defaults);


	$post_id = wp_insert_post($postarr, $wp_error = false );	

	if($post_category[1]!="") {wp_set_object_terms( $post_id, $post_category[1], $post_category[0] );}



unset($defaults);
unset($post_id);
unset($postarr);

$alert .= "<p>Post <b>$title</b> is created </p>";
    }
}
$alert .= "<br /><br /><a href='" . admin_url('edit.php?post_type=') . $post_type . "' target='_blank'>See posts</a> ";
}

?>


<style>

#d2p_table {
	background:#F9F9F9; 
	width:90%;
	padding:20px;
	font-size:14px;
}
#d2p_table td {
	padding:10px;
}
#d2p_results {
	background:#F9F9F9; 
	width:90%;
	padding:20px;

}
#d2p_results p {
 padding:4px 8px; 
 border-bottom:1px solid #F1F1F1;
 font-size:14px;
}
#d2p_results p::before {
	content: "\2713  ";
	color:green;
}
</style>


<div class="wrap">
<script>            

jQuery(document).ready(function($) {
 


$(function(){
    var conditionalSelect = $("#d2p_category"),
        // Save possible options
        options = conditionalSelect.children(".conditional").clone();
    
    $("#d2p_post_type").change(function(){
        var value = $(this).val();                  
        conditionalSelect.children(".conditional").remove();
        options.clone().filter("."+value).appendTo(conditionalSelect);
    }).trigger("change");
});


});



</script>
    <h2>Documents to Post</h2>

    <p>
        Create posts from Word documents or text files
    </p>
    <div id="d2p_results">
	<?php echo $alert; ?><br /><br />
	</div>
    <h3>Step by step guide</h3>
    
    <ol>
        <li>Select files and upload it, filename will be title and file content will be post content</li>
        <li>Choose post type, post, page or custom post type</li>
        <li>Choose right category from dropdown menu</li>
        <li>Txt and docx file extension is supported</li>
    </ol>
    
    

	    <form action="" method="post" enctype="multipart/form-data">
<table id="d2p_table">
<tr>
	<td width="200">
		Select documents to upload:<br />
		<i>(you can select multiple files)</i></td>
    <td><input type="file" name="fileToUpload[]" multiple id="fileToUpload"></td>
</tr>

<tr>
	<td>Choose post type: </td>

	<td>
<select name="type" id="d2p_post_type">
                        <option value="post">Post</option>
                        <option value="page">Page</option>
<?php
$class  = array();
$class[]  = ['cat' => 'category', 'css' => 'post'];
$taxonomies  = array();

$args = array(
   'public'   => true,
   '_builtin' => false
);
$output = 'objects';
$operator = 'and'; 

$post_types = get_post_types( $args, $output, $operator);
$taxonomies[] = "category";


	foreach ( $post_types  as $post_type ) {
  $taxonomy_objects = get_object_taxonomies( $post_type->name );
  if($taxonomy_objects[0]) {
  $taxonomies[] = $taxonomy_objects[0];
$class[] = ['cat' => $taxonomy_objects[0], 'css' => $post_type->name];
  }
echo '                        <option value="' . $post_type->name . '">' . $post_type->label . "</option>\n";
   
   
   
   }

?>
	</select>
	</td>
</tr>
<tr>
	<td>Choose category: </td>
	<td><select name="category" id="d2p_category">
                       <option value="0" selected="selected">Choose...</option>
<?php 

	$defaults = array( 'taxonomy' => $taxonomies, 'hide_empty'       => 0 );
    $args = wp_parse_args( $args, $defaults );


$categories =	get_categories($args);
foreach( $categories as $category ) {
echo '                       <option class="conditional ';
$i=0;

//if($category->taxonomy != 'category') {echo "c2p_hide ";}
foreach( $taxonomies as $taxo ) {
 if($category->taxonomy == $class[$i]['cat']) {echo $class[$i]['css'];}
$i++;
}
echo '" value="'.$category->taxonomy.'||'.$category->slug.'">'.$category->name.'</option>'."\n";

};


	?>



                    </select>
	</td>


</tr>
        <tr>
            <td>&nbsp;</td><td><input type="submit" name="create" value="Create" /><td>
</tr>
</table>
    </form>

</div>
