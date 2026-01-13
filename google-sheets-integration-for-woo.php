<?php
/**
 * Plugin Name: Google Sheets Integration for WooCommerce by WP Methods
 * Plugin URI: https://wpmethods.com/plugins/google-sheets-integration-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 2.0.5
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 * License: GPL2
 * Text Domain: wpmethods-wc-to-gs
 * Domain Path: /languages
 * Requires at least: 5.9
 * Requires PHP: 7.4
 */

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

define('GS_WC_VERSION', '2.0.5');

class WPMethods_WC_To_Google_Sheets {
    
    private $available_fields = array();
    
    public function __construct() {
        // Define available fields
        $this->define_available_fields();
        
        // Check if WooCommerce is active
        add_action('admin_init', array($this, 'wpmethods_check_woocommerce'));
        
        // Hook into order status changes
        add_action('woocommerce_order_status_changed', array($this, 'wpmethods_send_order_to_sheets'), 10, 4);
        
        // Add settings page
        add_action('admin_menu', array($this, 'wpmethods_add_admin_menu'));
        add_action('admin_init', array($this, 'wpmethods_settings_init'));
        
        // Add admin scripts
        add_action('admin_enqueue_scripts', array($this, 'wpmethods_admin_scripts'));
        
        // AJAX handler for generating Google Apps Script
        add_action('wp_ajax_wpmethods_generate_google_script', array($this, 'wpmethods_generate_google_script_ajax'));
        
        // Register activation hook to set default values
        register_activation_hook(__FILE__, array($this, 'wpmethods_activate_plugin'));
    }
    
    /**
     * Define available fields
     */
    private function define_available_fields() {
        // Simplified fields - only include working ones
        $this->available_fields = array(
            'order_id' => array(
                'label' => 'Order ID',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-cart'
            ),
            'billing_name' => array(
                'label' => 'Billing Name',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-admin-users'
            ),
            'billing_email' => array(
                'label' => 'Email Address',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-email'
            ),
            'billing_phone' => array(
                'label' => 'Phone',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-phone'
            ),
            'billing_address' => array(
                'label' => 'Billing Address',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-location'
            ),
            'product_name' => array(
                'label' => 'Product Name',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-products'
            ),
            'order_amount_with_currency' => array(
                'label' => 'Order Amount',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-money'
            ),
            'order_currency' => array(
                'label' => 'Order Currency',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-tag'
            ),
            'payment_method_title' => array(
                'label' => 'Payment Method',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-bank'
            ),
            'order_status' => array(
                'label' => 'Order Status',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-yes'
            ),
            'order_date' => array(
                'label' => 'Order Date',
                'required' => true,
                'always_include' => true,
                'icon' => 'dashicons dashicons-calendar'
            ),
            'product_categories' => array(
                'label' => 'Product Categories',
                'required' => false,
                'always_include' => false,
                'icon' => 'dashicons dashicons-category'
            )
        );
    }
    
    /**
     * Get selected fields for Google Sheets
     */
    private function wpmethods_get_selected_fields() {
        $selected_fields = get_option('wpmethods_wc_gs_selected_fields', array());
        
        // Ensure it's always an array
        if (!is_array($selected_fields)) {
            if (is_string($selected_fields) && !empty($selected_fields)) {
                $unserialized = maybe_unserialize($selected_fields);
                if (is_array($unserialized)) {
                    $selected_fields = $unserialized;
                } else {
                    $selected_fields = array_map('trim', explode(',', $selected_fields));
                }
            } else {
                $selected_fields = array();
            }
        }
        
        // Always include required fields
        foreach ($this->available_fields as $field_key => $field_info) {
            if (isset($field_info['always_include']) && $field_info['always_include']) {
                if (!in_array($field_key, $selected_fields)) {
                    $selected_fields[] = $field_key;
                }
            }
        }
        
        return array_unique($selected_fields);
    }
    
    /**
     * Check if WooCommerce is active
     */
    public function wpmethods_check_woocommerce() {
        if (!class_exists('WooCommerce')) {
            add_action('admin_notices', array($this, 'wpmethods_woocommerce_missing_notice'));
        }
    }
    
    /**
     * WooCommerce missing notice
     */
    public function wpmethods_woocommerce_missing_notice() {
        ?>
        <div class="notice notice-error">
            <p><?php _e('WP Methods WooCommerce to Google Sheets requires WooCommerce to be installed and activated.', 'wpmethods-wc-to-gs'); ?></p>
        </div>
        <?php
    }
    
    /**
     * Plugin activation - set default values
     */
    public function wpmethods_activate_plugin() {
        if (!get_option('wpmethods_wc_gs_order_statuses')) {
            update_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        }
        
        if (!get_option('wpmethods_wc_gs_script_url')) {
            update_option('wpmethods_wc_gs_script_url', '');
        }
        
        if (!get_option('wpmethods_wc_gs_product_categories')) {
            update_option('wpmethods_wc_gs_product_categories', array());
        }
        
        // Set default selected fields (all fields)
        if (!get_option('wpmethods_wc_gs_selected_fields')) {
            $default_fields = array_keys($this->available_fields);
            update_option('wpmethods_wc_gs_selected_fields', $default_fields);
        }
    }
    
    /**
     * Get all WooCommerce order statuses
     */
    private function wpmethods_get_wc_order_statuses() {
        $statuses = wc_get_order_statuses();
        $clean_statuses = array();
        
        foreach ($statuses as $key => $label) {
            $clean_key = str_replace('wc-', '', $key);
            $clean_statuses[$clean_key] = $label;
        }
        
        return $clean_statuses;
    }
    
    /**
     * Get selected categories as array
     */
    private function wpmethods_get_selected_categories() {
        $selected_categories = get_option('wpmethods_wc_gs_product_categories', array());
        
        if (!is_array($selected_categories)) {
            if (is_string($selected_categories) && !empty($selected_categories)) {
                $unserialized = maybe_unserialize($selected_categories);
                if (is_array($unserialized)) {
                    $selected_categories = $unserialized;
                } else {
                    $selected_categories = array_map('trim', explode(',', $selected_categories));
                }
            } else {
                $selected_categories = array();
            }
        }
        
        $selected_categories = array_map('intval', $selected_categories);
        
        return $selected_categories;
    }
    
    /**
     * Check if order contains products from selected categories
     */
    private function wpmethods_order_has_selected_categories($order, $selected_categories) {
        if (empty($selected_categories)) {
            return true;
        }
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                $product_categories = wp_get_post_terms($product->get_id(), 'product_cat', array('fields' => 'ids'));
                
                if (!empty(array_intersect($product_categories, $selected_categories))) {
                    return true;
                }
            }
        }
        
        return false;
    }
    
    /**
     * Send order data to Google Sheets
     */
    public function wpmethods_send_order_to_sheets($order_id, $old_status, $new_status, $order) {
        
        $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        
        if (!is_array($selected_statuses)) {
            $selected_statuses = maybe_unserialize($selected_statuses);
            if (!is_array($selected_statuses)) {
                $selected_statuses = array('completed', 'processing');
            }
        }
        
        if (!in_array($new_status, $selected_statuses)) {
            return;
        }
        
        $selected_categories = $this->wpmethods_get_selected_categories();
        
        if (!$this->wpmethods_order_has_selected_categories($order, $selected_categories)) {
            return;
        }
        
        $script_url = get_option('wpmethods_wc_gs_script_url', '');
        
        if (empty($script_url)) {
            error_log('WP Methods Google Sheets: Google Apps Script URL not configured');
            return;
        }
        
        $order_data = $this->wpmethods_prepare_order_data($order);
        
        $response = wp_remote_post($script_url, array(
            'method' => 'POST',
            'timeout' => 45,
            'redirection' => 5,
            'httpversion' => '1.0',
            'blocking' => true,
            'headers' => array(
                'Content-Type' => 'application/json',
            ),
            'body' => json_encode($order_data),
            'cookies' => array()
        ));
        
        if (is_wp_error($response)) {
            error_log('WP Methods Google Sheets integration error: ' . $response->get_error_message());
        } else {
            // Debug logging for successful submissions
            error_log('WP Methods Google Sheets: Order ' . $order_id . ' sent successfully');
        }
    }
    
    /**
     * Prepare order data for Google Sheets based on selected fields
     */
    private function wpmethods_prepare_order_data($order) {
        $selected_fields = $this->wpmethods_get_selected_fields();
        $order_data = array();
        
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $value = $this->get_field_value($field_key, $order);
                if ($value !== null) {
                    $order_data[$field_key] = $value;
                }
            }
        }
        
        return $order_data;
    }
    
    /**
     * Get value for a specific field
     */
    private function get_field_value($field_key, $order) {
        switch ($field_key) {
            case 'order_id':
                return $order->get_id();
                
            case 'billing_name':
                return $order->get_billing_first_name() . ' ' . $order->get_billing_last_name();
                
            case 'billing_email':
                return $order->get_billing_email();
                
            case 'billing_phone':
                return $order->get_billing_phone();
                
            case 'billing_address':
                $address_parts = array();
                
                if ($address1 = $order->get_billing_address_1()) {
                    $address_parts[] = $address1;
                }
                if ($address2 = $order->get_billing_address_2()) {
                    $address_parts[] = $address2;
                }
                if ($city = $order->get_billing_city()) {
                    $address_parts[] = $city;
                }
                if ($state = $order->get_billing_state()) {
                    $address_parts[] = $state;
                }
                if ($postcode = $order->get_billing_postcode()) {
                    $address_parts[] = $postcode;
                }
                if ($country = $order->get_billing_country()) {
                    $address_parts[] = $country;
                }
                
                return implode(', ', $address_parts);
                
            case 'product_name':
                $product_names = array();
                foreach ($order->get_items() as $item) {
                    $product_names[] = $item->get_name();
                }
                return implode(', ', $product_names);
                
            case 'order_amount_with_currency':
                $currency_symbol = $order->get_currency();
                $currency_symbol_formatted = get_woocommerce_currency_symbol($currency_symbol);
                // Fix: Don't use HTML entities, just send the raw symbol
                $symbol = $currency_symbol_formatted;
                // If it's an HTML entity, decode it
                if (strpos($symbol, '&#') !== false) {
                    $symbol = html_entity_decode($symbol, ENT_QUOTES, 'UTF-8');
                }
                return $symbol . $order->get_total();
                
            case 'order_currency':
                return $order->get_currency();
                
            case 'payment_method_title':
                $payment_title = $order->get_payment_method_title();
                if (empty($payment_title)) {
                    // Try to get payment gateway title
                    $payment_method = $order->get_payment_method();
                    if ($payment_method) {
                        $gateways = WC()->payment_gateways()->payment_gateways();
                        if (isset($gateways[$payment_method])) {
                            $payment_title = $gateways[$payment_method]->get_title();
                        }
                    }
                }
                return $payment_title ? $payment_title : '';
                
            case 'order_status':
                return $order->get_status();
                
            case 'order_date':
                $date_created = $order->get_date_created();
                return $date_created ? $date_created->format('Y-m-d H:i:s') : '';
                
            case 'product_categories':
                return $this->wpmethods_get_order_categories($order);
                
            default:
                return null;
        }
    }
    
    /**
     * Get categories from order products
     */
    private function wpmethods_get_order_categories($order) {
        $all_categories = array();
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                $categories = wp_get_post_terms($product->get_id(), 'product_cat', array('fields' => 'names'));
                $all_categories = array_merge($all_categories, $categories);
            }
        }
        
        $all_categories = array_unique($all_categories);
        
        return implode(', ', $all_categories);
    }
    
    /**
     * Add admin menu
     */
    public function wpmethods_add_admin_menu() {
        add_options_page(
            'WooCommerce to Google Sheets',
            'WC to Google Sheets',
            'manage_options',
            'wpmethods-wc-to-google-sheets',
            array($this, 'wpmethods_settings_page')
        );
    }
    
    /**
     * Admin scripts
     */
    public function wpmethods_admin_scripts($hook) {
        if ($hook != 'settings_page_wpmethods-wc-to-google-sheets') {
            return;
        }

        $min = defined('SCRIPT_DEBUG') && SCRIPT_DEBUG ? '' : '.min';
        
        // Enqueue WordPress core styles
        wp_enqueue_style('wp-color-picker');
        wp_enqueue_script('wp-color-picker');
        wp_enqueue_style('wpmethods-wc-gs-admin-style', plugin_dir_url(__FILE__) . 'assets/css/style' . $min . '.css', array(), GS_WC_VERSION);
        wp_enqueue_script('wpmethods-wc-gs-admin-script', plugin_dir_url(__FILE__) . 'assets/js/admin' . $min . '.js', array('jquery'), GS_WC_VERSION, true);
       
        // Localize script to pass PHP variables to JS
        wp_localize_script('wpmethods-wc-gs-admin-script', 'wpmethods_wc_gs', array(
            'ajax_url' => admin_url('admin-ajax.php'),
            'nonce' => wp_create_nonce('wpmethods_generate_script_nonce')
        ));
    }
    
    /**
     * Initialize settings
     */
    public function wpmethods_settings_init() {
        // Main settings
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_order_statuses', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_script_url', 'esc_url_raw');
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_product_categories', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_selected_fields', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_settings', 'wpmethods_wc_gs_monthly_sheets', array($this, 'wpmethods_sanitize_checkbox'));
        
        // Main settings section
        add_settings_section(
            'wpmethods_wc_gs_section',
            'Google Sheets Integration Settings',
            array($this, 'wpmethods_section_callback'),
            'wpmethods_wc_gs_settings'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_order_statuses',
            'Trigger Order Statuses',
            array($this, 'wpmethods_order_statuses_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_monthly_sheets',
            'Monthly Sheets',
            array($this, 'wpmethods_monthly_sheets_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_product_categories',
            'Product Categories Filter',
            array($this, 'wpmethods_product_categories_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );

        add_settings_field(
            'wpmethods_wc_gs_selected_fields',
            'Checkout Fields',
            array($this, 'wpmethods_selected_fields_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_script_url',
            'Google Apps Script URL',
            array($this, 'wpmethods_script_url_render'),
            'wpmethods_wc_gs_settings',
            'wpmethods_wc_gs_section'
        );
    }
    
    /**
     * Sanitize array inputs
     */
    public function wpmethods_sanitize_array($input) {
        if (!is_array($input)) {
            return array();
        }
        return array_map('sanitize_text_field', $input);
    }
    
    /**
     * Sanitize checkbox input
     */
    public function wpmethods_sanitize_checkbox($input) {
        return $input ? '1' : '0';
    }
    
    /**
     * Section callback
     */
    public function wpmethods_section_callback() {
        echo '<p>Configure the settings for Google Sheets integration.</p>';
    }
    
    /**
     * Monthly sheets field render
     */
    public function wpmethods_monthly_sheets_render() {
        $value = get_option('wpmethods_wc_gs_monthly_sheets', '0');
        ?>
        <label style="display: inline-flex; align-items: center; gap: 10px; padding: 15px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
            <input type="checkbox" name="wpmethods_wc_gs_monthly_sheets" value="1" <?php checked($value, '1'); ?> style="width: 20px; height: 20px;">
            <span style="font-weight: 500; font-size: 14px;">
                <span class="dashicons dashicons-calendar-alt" style="color: #667eea;"></span>
                Enable automatic monthly sheet creation
            </span>
        </label>
        <p class="description" style="margin-top: 10px;">When enabled, a new sheet will be automatically created for each month (e.g., "January 2024", "February 2024")</p>
        <?php
    }
    
    /**
     * Order Statuses field render
     */
    public function wpmethods_order_statuses_render() {
        $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array('completed', 'processing'));
        
        if (!is_array($selected_statuses)) {
            $selected_statuses = maybe_unserialize($selected_statuses);
            if (!is_array($selected_statuses)) {
                $selected_statuses = array('completed', 'processing');
            }
        }
        
        $all_statuses = $this->wpmethods_get_wc_order_statuses();
        
        echo '<div style="display: grid; grid-template-columns: repeat(auto-fill, minmax(200px, 1fr)); gap: 10px; margin-bottom: 15px;">';
        
        foreach ($all_statuses as $status => $label) {
            $checked = in_array($status, $selected_statuses) ? 'checked' : '';
            ?>
            <div style="padding: 12px; background: #f8f9fa; border-radius: 6px; border: 1px solid #e0e0e0;">
                <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                    <input type="checkbox" name="wpmethods_wc_gs_order_statuses[]" 
                           value="<?php echo esc_attr($status); ?>" <?php echo $checked; ?> style="width: 18px; height: 18px;">
                    <span style="font-weight: 500; font-size: 14px;"><?php echo esc_html($label); ?></span>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description" style="margin-top: 10px;">Select order statuses that should trigger sending data to Google Sheets</p>';
    }
    
    /**
     * Get all WooCommerce product categories
     */
    private function wpmethods_get_product_categories() {
        $categories = get_terms(array(
            'taxonomy' => 'product_cat',
            'hide_empty' => false,
            'orderby' => 'name',
            'order' => 'ASC',
        ));
        
        $category_list = array();
        
        if (!is_wp_error($categories) && !empty($categories)) {
            foreach ($categories as $category) {
                $category_list[$category->term_id] = $category->name;
            }
        }
        
        return $category_list;
    }
    
    /**
     * Product Categories field render
     */
    public function wpmethods_product_categories_render() {
        $selected_categories = $this->wpmethods_get_selected_categories();
        $all_categories = $this->wpmethods_get_product_categories();
        
        if (empty($all_categories)) {
            echo '<p>No product categories found.</p>';
            return;
        }
        
        echo '<div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 15px; margin-bottom: 10px; background: #f8f9fa; border-radius: 6px;">';
        
        foreach ($all_categories as $cat_id => $cat_name) {
            $checked = in_array($cat_id, $selected_categories) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 8px; padding: 10px; background: white; border-radius: 4px; border: 1px solid #e0e0e0;">
                <label style="display: flex; align-items: center; gap: 10px; cursor: pointer;">
                    <input type="checkbox" name="wpmethods_wc_gs_product_categories[]" 
                           value="<?php echo esc_attr($cat_id); ?>" <?php echo $checked; ?> style="width: 18px; height: 18px;">
                    <span style="font-weight: 500; font-size: 14px;"><?php echo esc_html($cat_name); ?></span>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description" style="margin-top: 10px;">Select product categories. Orders will only be sent if they contain at least one product from selected categories. Leave empty to include all categories.</p>';
    }
    
    /**
     * Selected fields render
     */
    public function wpmethods_selected_fields_render() {
        $selected_fields = $this->wpmethods_get_selected_fields();
        
        // Search box
        ?>
        <div class="wpmethods-search-box">
            <input type="text" id="wpmethods-field-search" placeholder="Search fields...">
        </div>
        <?php
        
        // Display all fields in categories
        echo '<div class="wpmethods-field-category">';
        echo '<h3>';
        echo '<span><span class="dashicons dashicons-list-view"></span> All Fields</span>';
        echo '<span class="dashicons dashicons-arrow-down"></span>';
        echo '</h3>';
        echo '<div class="wpmethods-fields-grid">';
        
        foreach ($this->available_fields as $field_key => $field_info) {
            $checked = in_array($field_key, $selected_fields) ? 'checked' : '';
            $disabled = (isset($field_info['always_include']) && $field_info['always_include']) ? 'disabled' : '';
            $required = (isset($field_info['required']) && $field_info['required']) ? ' <span class="required">Required</span>' : '';
            $icon = isset($field_info['icon']) ? $field_info['icon'] : 'dashicons dashicons-admin-generic';
            ?>
            <div class="wpmethods-field-item">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_selected_fields[]" 
                           value="<?php echo esc_attr($field_key); ?>" 
                           <?php echo $checked; ?> <?php echo $disabled; ?>>
                    <span class="<?php echo esc_attr($icon); ?>" style="color: #667eea;"></span>
                    <span><?php echo esc_html($field_info['label']); ?></span>
                    <?php echo $required; ?>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '</div>';
        
        echo '<p class="description" style="margin-top: 20px;">Select fields to include in Google Sheets. Required fields are always included and cannot be disabled.</p>';
        ?>
        <div class="wpmethods-feature-box">
            <h3 style="margin-top: 0; color: #667eea;">
                <span class="dashicons dashicons-code-standards"></span> Generate Google Apps Script
            </h3>
            <p>Click the button below to generate a Google Apps Script code based on your selected fields. This script will handle data submission to your Google Sheets.</p>
            <button type="button" id="wpmethods-generate-script" class="wpmethods-button">
                <span class="dashicons dashicons-update"></span> Generate Google Apps Script
            </button>
            <div id="wpmethods-script-output" style="margin-top: 15px; display: none;">
                <h4>Generated Script</h4>
                <textarea id="wpmethods-generated-script" style="width: 100%; height: 400px;" readonly></textarea>
                <p style="margin-top: 10px;">
                    <button type="button" id="wpmethods-copy-script" class="wpmethods-button wpmethods-button-secondary">
                        <span class="dashicons dashicons-clipboard"></span> Copy to Clipboard
                    </button>
                    <span id="wpmethods-copy-status" style="margin-left: 10px; color: #28a745; font-weight: 600; display: none;">
                        <span class="dashicons dashicons-yes-alt"></span> Copied!
                    </span>
                </p>
            </div>
        </div>
        <?php
    }
    
    /**
     * Script URL field render
     */
    public function wpmethods_script_url_render() {
        $value = get_option('wpmethods_wc_gs_script_url', '');
        ?>
        <div style="max-width: 600px;">
            <input type="url" name="wpmethods_wc_gs_script_url" 
                   value="<?php echo esc_url($value); ?>" 
                   style="width: 100%; padding: 12px; border: 2px solid #e0e0e0; border-radius: 6px; font-size: 14px;" 
                   placeholder="https://script.google.com/macros/s/...">
            <p class="description" style="margin-top: 10px;">
                <span class="dashicons dashicons-info" style="color: #667eea;"></span>
                Enter your Google Apps Script web app URL. Get this from Google Apps Script deployment after copying the generated script.
            </p>
        </div>
        <?php
    }
    
    /**
     * AJAX handler for generating Google Apps Script
     */
    public function wpmethods_generate_google_script_ajax() {
        check_ajax_referer('wpmethods_generate_script_nonce', 'nonce');
        
        if (!current_user_can('manage_options')) {
            wp_die('Unauthorized');
        }
        
        $selected_fields = isset($_POST['fields']) ? (array) $_POST['fields'] : $this->wpmethods_get_selected_fields();
        
        // Check if monthly sheets option is enabled
        $monthly_sheets = get_option('wpmethods_wc_gs_monthly_sheets', '0');
        
        // Generate Google Apps Script code
        if ($monthly_sheets === '1') {
            $script = $this->generate_google_apps_script_monthly($selected_fields);
        } else {
            $script = $this->generate_google_apps_script_single($selected_fields);
        }
        
        wp_send_json_success(array(
            'script' => $script,
            'fields' => $selected_fields,
            'monthly_sheets' => $monthly_sheets
        ));
    }
    
    /**
     * Generate Google Apps Script code for single sheet
     */
    private function generate_google_apps_script_single($selected_fields) {
        // Get field labels for headers
        $headers = array();
        $field_mapping = array();
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $headers[] = $this->available_fields[$field_key]['label'];
                $field_mapping[$field_key] = $this->available_fields[$field_key]['label'];
            }
        }
        
        $headers_js = json_encode($headers, JSON_PRETTY_PRINT);
        $field_mapping_js = json_encode($field_mapping, JSON_PRETTY_PRINT);
        
        $script = <<<EOT
// Google Apps Script Code for Google Sheets
// Generated by WP Methods WooCommerce to Google Sheets Plugin
// Single Sheet Mode
// Fields: {$this->get_field_list($selected_fields)}

function doPost(e) {
    try {
        // Parse the incoming data
        const data = JSON.parse(e.postData.contents);
        
        // Get the active sheet
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
        
        // Initialize headers if sheet is empty
        initializeSheet(sheet);
        
        // Check if this order already exists
        const orderIds = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
        const existingRowIndex = orderIds.indexOf(data.order_id.toString());
        
        if (existingRowIndex !== -1) {
            // Update existing row
            updateExistingRow(sheet, existingRowIndex, data);
        } else {
            // Add new row
            addNewRow(sheet, data);
        }
        
        // Return success response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Order data saved successfully'
        })).setMimeType(ContentService.MimeType.JSON);
        
    } catch (error) {
        // Return error response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function initializeSheet(sheet) {
    if (sheet.getLastRow() === 0) {
        const headers = {$headers_js};
        sheet.appendRow(headers);
        
        // Format header row
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#4CAF50')
                   .setFontColor('white')
                   .setFontWeight('bold');
        
        // Set column widths
        for (let i = 1; i <= headers.length; i++) {
            sheet.autoResizeColumn(i);
        }
        
        // Freeze header row
        sheet.setFrozenRows(1);
    }
}

function updateExistingRow(sheet, existingRowIndex, data) {
    const row = existingRowIndex + 2; // +2 for header row and 0-based index
    
    const fieldOrder = {$headers_js};
    
    fieldOrder.forEach((fieldLabel, index) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        if (data[fieldKey] !== undefined) {
            sheet.getRange(row, index + 1).setValue(data[fieldKey]);
        }
    });
}

function addNewRow(sheet, data) {
    const fieldOrder = {$headers_js};
    const rowData = [];
    
    fieldOrder.forEach((fieldLabel) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        rowData.push(data[fieldKey] || '');
    });
    
    sheet.appendRow(rowData);
    
    // Apply alternating row colors for readability
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const rowRange = sheet.getRange(lastRow, 1, 1, fieldOrder.length);
        if (lastRow % 2 === 0) {
            rowRange.setBackground('#F5F5F5'); // Light gray for even rows
        }
    }
}

function getFieldKeyFromLabel(fieldLabel) {
    const fieldMap = {$field_mapping_js};
    
    // Reverse lookup: find key by label
    for (const [key, label] of Object.entries(fieldMap)) {
        if (label === fieldLabel) {
            return key;
        }
    }
    
    // Fallback: convert label to lowercase with underscores
    return fieldLabel.toLowerCase().replace(/ /g, '_');
}

// Function to manually initialize the sheet with headers
function manualInitialize() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    initializeSheet(sheet);
}
EOT;

        return $script;
    }
    
    /**
     * Generate Google Apps Script code for monthly sheets
     */
    private function generate_google_apps_script_monthly($selected_fields) {
        // Get field labels for headers
        $headers = array();
        $field_mapping = array();
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $headers[] = $this->available_fields[$field_key]['label'];
                $field_mapping[$field_key] = $this->available_fields[$field_key]['label'];
            }
        }
        
        $headers_js = json_encode($headers, JSON_PRETTY_PRINT);
        $field_mapping_js = json_encode($field_mapping, JSON_PRETTY_PRINT);
        
        $script = <<<EOT
// Google Apps Script Code for Google Sheets
// Generated by WP Methods WooCommerce to Google Sheets Plugin
// Monthly Sheets Mode - Automatically creates new sheet for each month
// Fields: {$this->get_field_list($selected_fields)}

function doPost(e) {
    try {
        // Parse the incoming data
        const data = JSON.parse(e.postData.contents);
        
        // Get month from order date (format: YYYY-MM-DD HH:MM:SS)
        const orderDate = data.order_date;
        let monthYear = getMonthYearFromDate(orderDate);
        
        // Get or create sheet for this month
        const sheet = getOrCreateMonthlySheet(monthYear);
        
        // Check if this order already exists in this month's sheet
        const orderIds = sheet.getRange(2, 1, sheet.getLastRow(), 1).getValues().flat();
        const existingRowIndex = orderIds.indexOf(data.order_id.toString());
        
        if (existingRowIndex !== -1) {
            // Update existing row
            updateExistingRow(sheet, existingRowIndex, data);
        } else {
            // Add new row
            addNewRow(sheet, data);
        }
        
        // Return success response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'success',
            message: 'Order data saved to ' + monthYear + ' sheet successfully'
        })).setMimeType(ContentService.MimeType.JSON);
        
    } catch (error) {
        // Return error response
        return ContentService.createTextOutput(JSON.stringify({
            status: 'error',
            message: error.toString()
        })).setMimeType(ContentService.MimeType.JSON);
    }
}

function getMonthYearFromDate(dateString) {
    // Parse date string (format: YYYY-MM-DD HH:MM:SS)
    const dateParts = dateString.split(' ')[0].split('-');
    const year = dateParts[0];
    const monthNum = parseInt(dateParts[1]);
    
    // Month names
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    const monthName = monthNames[monthNum - 1];
    return monthName + ' ' + year;
}

function getOrCreateMonthlySheet(monthYear) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(monthYear);
    
    if (!sheet) {
        // Create new sheet for the month
        sheet = spreadsheet.insertSheet(monthYear);
        
        // Add headers to the new sheet
        const headers = {$headers_js};
        sheet.appendRow(headers);
        
        // Format header row
        const headerRange = sheet.getRange(1, 1, 1, headers.length);
        headerRange.setBackground('#4CAF50')
                   .setFontColor('white')
                   .setFontWeight('bold');
        
        // Set column widths
        for (let i = 1; i <= headers.length; i++) {
            sheet.autoResizeColumn(i);
        }
        
        // Freeze header row
        sheet.setFrozenRows(1);
        
        // Log sheet creation
        console.log('Created new sheet: ' + monthYear);
    }
    
    return sheet;
}

function updateExistingRow(sheet, existingRowIndex, data) {
    const row = existingRowIndex + 2; // +2 for header row and 0-based index
    
    const fieldOrder = {$headers_js};
    
    fieldOrder.forEach((fieldLabel, index) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        if (data[fieldKey] !== undefined) {
            sheet.getRange(row, index + 1).setValue(data[fieldKey]);
        }
    });
    
    // Highlight updated row
    const rowRange = sheet.getRange(row, 1, 1, fieldOrder.length);
    rowRange.setBackground('#FFF9C4'); // Light yellow background
}

function addNewRow(sheet, data) {
    const fieldOrder = {$headers_js};
    const rowData = [];
    
    fieldOrder.forEach((fieldLabel) => {
        const fieldKey = getFieldKeyFromLabel(fieldLabel);
        rowData.push(data[fieldKey] || '');
    });
    
    sheet.appendRow(rowData);
    
    // Apply alternating row colors for readability
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const rowRange = sheet.getRange(lastRow, 1, 1, fieldOrder.length);
        if (lastRow % 2 === 0) {
            rowRange.setBackground('#F5F5F5'); // Light gray for even rows
        }
    }
}

function getFieldKeyFromLabel(fieldLabel) {
    const fieldMap = {$field_mapping_js};
    
    // Reverse lookup: find key by label
    for (const [key, label] of Object.entries(fieldMap)) {
        if (label === fieldLabel) {
            return key;
        }
    }
    
    // Fallback: convert label to lowercase with underscores
    return fieldLabel.toLowerCase().replace(/ /g, '_');
}

// Function to manually create sheets for past/future months
function createAllMonthlySheets() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // Create sheets for current year and next year
    const currentYear = new Date().getFullYear();
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    
    // Create sheets for current year
    for (let month = 0; month < 12; month++) {
        const monthYear = monthNames[month] + ' ' + currentYear;
        getOrCreateMonthlySheet(monthYear);
    }
    
    // Create sheets for next year
    const nextYear = currentYear + 1;
    for (let month = 0; month < 12; month++) {
        const monthYear = monthNames[month] + ' ' + nextYear;
        getOrCreateMonthlySheet(monthYear);
    }
    
    console.log('Created monthly sheets for ' + currentYear + ' and ' + nextYear);
}

// Function to initialize with current month sheet
function initializeCurrentMonth() {
    const currentDate = new Date();
    const monthNames = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ];
    const monthYear = monthNames[currentDate.getMonth()] + ' ' + currentDate.getFullYear();
    
    getOrCreateMonthlySheet(monthYear);
    console.log('Initialized current month sheet: ' + monthYear);
}
EOT;

        return $script;
    }
    
    /**
     * Get field list string
     */
    private function get_field_list($selected_fields) {
        $field_names = array();
        foreach ($selected_fields as $field_key) {
            if (isset($this->available_fields[$field_key])) {
                $field_names[] = $this->available_fields[$field_key]['label'];
            }
        }
        return implode(', ', $field_names);
    }
    
    /**
     * Settings page with modern design
     */
    public function wpmethods_settings_page() {
        $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array());
        $selected_fields = $this->wpmethods_get_selected_fields();
        $monthly_sheets = get_option('wpmethods_wc_gs_monthly_sheets', '0');
        $selected_categories = $this->wpmethods_get_selected_categories();
        ?>
        <div class="wrap wpmethods-settings-wrapper">
            
            <!-- Modern Header -->
            <div class="wpmethods-header">
                <h1>
                    <span class="dashicons dashicons-google" style="vertical-align: middle; margin-right: 10px;"></span>
                    WooCommerce to Google Sheets
                </h1>
                <p>Automatically send WooCommerce orders to Google Sheets. Configure your integration below.</p>
            </div>
            
            <!-- Dashboard Stats -->
            <div class="wpmethods-dashboard">
                <div class="wpmethods-card">
                    <h3><span class="dashicons dashicons-admin-settings"></span> Configuration Status</h3>
                    <div class="wpmethods-stats-grid">
                        <div class="wpmethods-stat">
                            <div class="wpmethods-stat-number"><?php echo count($selected_fields); ?></div>
                            <div class="wpmethods-stat-label">Selected Fields</div>
                        </div>
                        <div class="wpmethods-stat">
                            <div class="wpmethods-stat-number"><?php echo count($selected_statuses); ?></div>
                            <div class="wpmethods-stat-label">Trigger Statuses</div>
                        </div>
                        <div class="wpmethods-stat">
                            <div class="wpmethods-stat-number"><?php echo $monthly_sheets === '1' ? '✓' : '—'; ?></div>
                            <div class="wpmethods-stat-label">Monthly Sheets</div>
                        </div>
                    </div>
                </div>
                
                <div class="wpmethods-card">
                    <h3><span class="dashicons dashicons-lightbulb"></span> Quick Tips</h3>
                    <ul style="margin: 0; padding-left: 20px; color: #666;">
                        <li style="margin-bottom: 8px;">Required fields are automatically included</li>
                        <li style="margin-bottom: 8px;">Use monthly sheets for better organization</li>
                        <li style="margin-bottom: 8px;">Test with one order status first</li>
                        <li>Watch the setup video for detailed instructions</li>
                    </ul>
                </div>
            </div>
            
            <!-- Main Settings Form -->
            <form action="options.php" method="post" style="margin-bottom: 30px;">
                <?php
                settings_fields('wpmethods_wc_gs_settings');
                do_settings_sections('wpmethods_wc_gs_settings');
                submit_button('Save Settings', 'primary wpmethods-button', 'submit', false);
                ?>
            </form>
            
            <!-- Video Tutorial -->
            <div class="wpmethods-video-box">
                <h3><span class="dashicons dashicons-video-alt3"></span> Watch Setup Tutorial</h3>
                <p style="color: rgba(255,255,255,0.9); margin-bottom: 20px;">Learn how to set up the plugin step by step with our video tutorial.</p>
                <a href="https://youtu.be/7Kh-uugbods" target="_blank" class="wpmethods-video-button">
                    <span class="dashicons dashicons-youtube"></span> Watch Tutorial Video
                </a>
            </div>
            
            <!-- Donation Section -->
            <div class="wpmethods-donation-box">
                <h3><span class="dashicons dashicons-heart"></span> Support This Plugin</h3>
                <p style="color: #856404; margin-bottom: 20px;">If this plugin has helped your business, consider buying me a coffee to support further development.</p>
                <a href="https://buymeacoffee.com/ajharrashed" target="_blank" class="wpmethods-donation-button">
                    <span class="dashicons dashicons-coffee"></span> Buy Me a Coffee
                </a>
            </div>
            
            <!-- Configuration Summary -->
            <div class="wpmethods-card">
                <h3><span class="dashicons dashicons-chart-bar"></span> Configuration Summary</h3>
                
                <div class="wpmethods-stats-grid">
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number">
                            <?php 
                            if (!is_array($selected_statuses)) {
                                $selected_statuses = maybe_unserialize($selected_statuses);
                                if (!is_array($selected_statuses)) {
                                    $selected_statuses = array('completed', 'processing');
                                }
                            }
                            echo count($selected_statuses);
                            ?>
                        </div>
                        <div class="wpmethods-stat-label">Trigger Statuses</div>
                    </div>
                    
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number"><?php echo $monthly_sheets === '1' ? 'Enabled' : 'Single'; ?></div>
                        <div class="wpmethods-stat-label">Sheet Mode</div>
                    </div>
                    
                    <div class="wpmethods-stat">
                        <div class="wpmethods-stat-number"><?php echo count($selected_fields); ?></div>
                        <div class="wpmethods-stat-label">Selected Fields</div>
                    </div>
                </div>
                
                <div style="margin-top: 20px; padding: 15px; background: #f8f9fa; border-radius: 6px;">
                    <h4 style="margin-top: 0; margin-bottom: 10px;">Selected Fields:</h4>
                    <p style="margin: 0; color: #666;">
                        <?php 
                        $field_names = array();
                        foreach ($selected_fields as $field_key) {
                            if (isset($this->available_fields[$field_key])) {
                                $field_names[] = $this->available_fields[$field_key]['label'];
                            }
                        }
                        echo implode(', ', $field_names);
                        ?>
                    </p>
                </div>
                
                <?php if ($monthly_sheets === '1'): ?>
                    <div style="margin-top: 20px; padding: 15px; background: #e8f5e9; border-radius: 6px; border-left: 4px solid #4CAF50;">
                        <h4 style="margin-top: 0; color: #2e7d32;">
                            <span class="dashicons dashicons-calendar-alt"></span> Monthly Sheets Active
                        </h4>
                        <p style="margin: 10px 0 0 0; color: #2e7d32;">
                            Orders will be automatically organized into monthly sheets (<?php echo date('F Y'); ?>, <?php echo date('F Y', strtotime('+1 month')); ?>, etc.)
                        </p>
                    </div>
                <?php endif; ?>
            </div>
            
        </div>
        <?php
    }
}

// Initialize the plugin
new WPMethods_WC_To_Google_Sheets();