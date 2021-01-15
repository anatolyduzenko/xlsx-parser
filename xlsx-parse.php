<?
/**
 * Plugin Name: Anatoly's Xlsx Parser
 * Version: 1.0.0
 * Author: Anatoly Duzenko
 * Author URI: www.hmmnuok.ru
 * Description: This Plugin used to parse Excel files and convert them into appropriate format (or something else)
 */

if ( ! defined( 'WPINC' ) ) {
	die;
}

require_once( plugin_dir_path( __FILE__ ) . 'vendor/autoload.php' );


add_action('admin_menu', 'xlsx_parser_plugin_setup_menu');
 
function xlsx_parser_plugin_setup_menu(){
    add_menu_page( 'Xlsx Processor', 'Xlsx Processor', 'manage_options', 'xlsx-parser', 'xlsx_parser_run_plugin' );
}
 
require plugin_dir_path( __FILE__ ) . 'includes/ADXlsxParser.php';

function xlsx_parser_run_plugin() {

	$plugin = new ADXlsxParser();
	$plugin->run();

}

function xlsx_parser_load_custom_wp_admin_style($hook) {
	if( $hook != 'toplevel_page_xlsx-parser' ) {
		 return;
	}
	wp_enqueue_style( 'custom_wp_admin_css', plugins_url('/admin/css/admin-style.css', __FILE__) );
}
add_action( 'admin_enqueue_scripts', 'xlsx_parser_load_custom_wp_admin_style' );

add_action( 'admin_post_nopriv_xlsx_parser_process_form_data', 'xlsx_parser_process_form_data' );
add_action( 'admin_post_xlsx_parser_process_form_data', 'xlsx_parser_process_form_data' );

function xlsx_parser_process_form_data() {
	$plugin = new ADXlsxParser();
	$plugin->analyze_file();
	wp_safe_redirect(admin_url('admin.php?page=xlsx-parser'));
	exit();
}


add_action( 'admin_post_nopriv_xlsx_parser_convert_file', 'xlsx_parser_convert_file' );
add_action( 'admin_post_xlsx_parser_convert_file', 'xlsx_parser_convert_file' );

function xlsx_parser_convert_file() {
	$plugin = new ADXlsxParser();
	$plugin->convert_file();
	// $plugin->run();
}

function xlsx_parser_start_session()
{
    if (!session_id())
        session_start();
}
add_action("init", "xlsx_parser_start_session", 1);

function xlsx_parser_upload_folder($upload) 
{
	$upload['subdir'] = "";
	$upload['path']   = plugin_dir_path( __FILE__ ).'files';
	$upload['url']    = "";
	return $upload;
}