<?

/**
 *
 * XLSX Parser Class 
 *
 */

class ADXlsxParser 
{

	private $post;
	
	function __construct()
	{
		
	}

	function run ()
	{
		$html =  "<h2>Welocme to Xlsx Parser!</h2><hr>";

		// File upload form
		$html .= "<form enctype='multipart/form-data' method='post' action=".admin_url( 'admin-post.php' ).">";
		$html .= "<input type='hidden' name='action' value='xlsx_parser_process_form_data'>";
		$html .= "<input type='file' name='xlsx_file' accept='.xls, .xlsx' required/>";
		$html .= "<input type='submit' name='submit' value='Analyze File'/>";
		$html .= "</form>";

		$html .= "<div id='parse-analysis-result'>".$_SESSION['result']."</div>";

		echo $html;
	}

	function analyze_file()
	{
		if ( ! function_exists( 'wp_handle_upload' ) ) {
		    require_once( ABSPATH . 'wp-admin/includes/file.php' );
		}
		if ( ! empty( $_POST ) ) {
			if(!empty($_FILES['xlsx_file'])) {
				$_SESSION['xlsx_parsed_file'] = $_FILES['xlsx_file']['name'];
				add_filter('upload_dir', 'xlsx_parser_upload_folder');
				$movefile = wp_handle_upload( $_FILES['xlsx_file'], ['test_form' => false] );
				remove_filter('upload_dir', 'xlsx_parser_upload_folder');
				if(isset($movefile['error'])) {
					$_SESSION['result'] = "<span class='error'>".$movefile['error']."</span>";
				} else {
					echo "File is valid, and was successfully uploaded.\n";
   				
					unset($_SESSION['result']);
				}
			} else {
				$_SESSION['result'] = "<span class='error'> The file is required.</span>";
			}
		}
	}
}