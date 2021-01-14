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

require plugin_dir_path( __FILE__ ) . 'includes/ADXlsxParser.php';

function run_plugin_name() {

	$plugin = new ADXlsxParser();
	$plugin->run();

}
run_plugin_name();