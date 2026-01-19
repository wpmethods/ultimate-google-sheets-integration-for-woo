<?php
/**
 * Plugin Name: Send Orders to Google Sheets for WooCommerce
 * Plugin URI: https://wpmethods.com/product/send-orders-to-google-sheets-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 1.0.0
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 * License: GPL v2 or later
 * License URI: https://www.gnu.org/licenses/gpl-2.0.html
 * Text Domain: send-orders-to-google-sheets-for-woocommerce
 * Domain Path: /languages
 * Requires at least: 5.9
 * Requires PHP: 7.4
 */

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}
if ( ! defined('UGSIW_VERSION') ) {
    define('UGSIW_VERSION', '1.0.0');
}


// Define plugin path and url
define('UGSIW_WPMETHODS_PATH', plugin_dir_path(__FILE__));
define('UGSIW_WPMETHODS_URL', plugin_dir_url(__FILE__));
// Include the autoloader
require_once UGSIW_WPMETHODS_PATH . 'includes/file-autoloader.php';

// Initialize
add_action('plugins_loaded', function() {
    new UGSIW\UGSIW_To_Google_Sheets();
});