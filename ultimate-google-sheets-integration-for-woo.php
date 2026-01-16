<?php
/**
 * Plugin Name: Ultimate Google Sheets Integration for WooCommerce
 * Plugin URI: https://wpmethods.com/plugins/google-sheets-integration-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 1.0.0
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 * License: GPL2
 * Text Domain: ultimate-google-sheets-integration-for-woo
 * Domain Path: /languages
 * Requires at least: 5.9
 * Requires PHP: 7.4
 * prefix: ugsiw
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
    if (is_admin()) {
        new UGSIW\UGSIW_To_Google_Sheets();
    } 
});