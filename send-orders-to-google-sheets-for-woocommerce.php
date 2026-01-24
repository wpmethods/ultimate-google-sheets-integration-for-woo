<?php
/**
 * Plugin Name: Send Orders to Google Sheets for WooCommerce
 * Plugin URI: https://wpmethods.com/product/send-orders-to-google-sheets-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 1.0.1
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 * License: GPL v2 or later
 * License URI: https://www.gnu.org/licenses/gpl-2.0.html
 * Text Domain: send-orders-to-google-sheets-for-woocommerce
 * Domain Path: /languages
 * Requires at least: 5.9
 * Requires PHP: 7.4
 */

use YahnisElsts\PluginUpdateChecker\v5\PucFactory;

if (!defined('ABSPATH')) {
    exit;
}

define('UGSIW_VERSION', '1.0.1');
define('UGSIW_PATH', plugin_dir_path(__FILE__));
define('UGSIW_URL', plugin_dir_url(__FILE__));

// Autoloader
require_once UGSIW_PATH . 'includes/file-autoloader.php';

// Load Plugin Update Checker
require_once UGSIW_PATH . 'updates/plugin-update-checker.php';

// Init plugin
add_action('plugins_loaded', function () {
    new UGSIW\UGSIW_To_Google_Sheets();
});

// Build update checker
$ugsiw_update_checker = PucFactory::buildUpdateChecker(
    'https://github.com/wpmethods/send-orders-to-google-sheets-for-woocommerce/',
    __FILE__,
    'send-orders-to-google-sheets-for-woocommerce'
);

// Stable branch
$ugsiw_update_checker->setBranch('main');
