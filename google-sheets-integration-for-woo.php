<?php
/**
 * Plugin Name: Google Sheets Integration for WooCommerce by WP Methods
 * Plugin URI: https://wpmethods.com/plugins/google-sheets-integration-for-woocommerce/
 * Description: Send order data to Google Sheets when order status changes to selected statuses
 * Version: 2.0.3
 * Author: WP Methods
 * Author URI: https://wpmethods.com
 */

// Prevent direct access
if (!defined('ABSPATH')) {
    exit;
}

class WPMethods_WC_To_Google_Sheets {
    
    private $available_fields = array();
    private $default_fields = array();
    
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
        // Default fields that are always available
        $this->default_fields = array(
            'order_id' => array(
                'label' => 'Order ID',
                'required' => true,
                'always_include' => true
            ),
            'billing_name' => array(
                'label' => 'Billing Name',
                'required' => true,
                'always_include' => true
            ),
            'billing_phone' => array(
                'label' => 'Billing Phone',
                'required' => false,
                'always_include' => false
            ),
            'billing_address' => array(
                'label' => 'Billing Address',
                'required' => false,
                'always_include' => false
            ),
            'product_name' => array(
                'label' => 'Product Name',
                'required' => true,
                'always_include' => true
            ),
            'order_amount' => array(
                'label' => 'Order Amount',
                'required' => true,
                'always_include' => true
            ),
            'order_status' => array(
                'label' => 'Order Status',
                'required' => true,
                'always_include' => true
            ),
            'order_date' => array(
                'label' => 'Order Date',
                'required' => true,
                'always_include' => true
            ),
            'product_categories' => array(
                'label' => 'Product Categories',
                'required' => false,
                'always_include' => false
            ),
            'shipping_class' => array(
                'label' => 'Shipping Class',
                'required' => false,
                'always_include' => false
            )
        );
        
        // Start with default fields
        $this->available_fields = $this->default_fields;
        
        // Add checkout fields dynamically when needed
        add_action('admin_init', array($this, 'wpmethods_add_checkout_fields'), 20);
    }
    
    /**
     * Add WooCommerce checkout fields (called later in admin_init)
     */
    public function wpmethods_add_checkout_fields() {
        if (!class_exists('WooCommerce')) {
            return;
        }
        
        $checkout_fields = $this->get_wc_checkout_fields();
        
        // Add checkout fields to available fields
        foreach ($checkout_fields as $key => $field) {
            $this->available_fields[$key] = $field;
        }
    }
    
    /**
     * Get WooCommerce checkout fields
     */
    private function get_wc_checkout_fields() {
        $checkout_fields = array();
        
        // Get billing fields from WooCommerce countries class
        if (class_exists('WC_Countries')) {
            $countries = new WC_Countries();
            
            // Billing fields
            $billing_fields = $countries->get_address_fields('', 'billing_');
            foreach ($billing_fields as $key => $field) {
                $clean_key = str_replace('billing_', '', $key);
                $checkout_fields['billing_' . $clean_key] = array(
                    'label' => isset($field['label']) ? $field['label'] : ucfirst(str_replace('_', ' ', $clean_key)),
                    'required' => isset($field['required']) ? $field['required'] : false,
                    'always_include' => false,
                    'type' => 'billing',
                    'original_key' => $key
                );
            }
            
            // Shipping fields
            $shipping_fields = $countries->get_address_fields('', 'shipping_');
            foreach ($shipping_fields as $key => $field) {
                $clean_key = str_replace('shipping_', '', $key);
                $checkout_fields['shipping_' . $clean_key] = array(
                    'label' => isset($field['label']) ? $field['label'] : ucfirst(str_replace('_', ' ', $clean_key)),
                    'required' => isset($field['required']) ? $field['required'] : false,
                    'always_include' => false,
                    'type' => 'shipping',
                    'original_key' => $key
                );
            }
        }
        
        // Additional order fields
        $additional_fields = array(
            'customer_note' => array(
                'label' => 'Customer Note',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'payment_method' => array(
                'label' => 'Payment Method',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'payment_method_title' => array(
                'label' => 'Payment Method Title',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'transaction_id' => array(
                'label' => 'Transaction ID',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'customer_ip' => array(
                'label' => 'Customer IP',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'customer_user_agent' => array(
                'label' => 'User Agent',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'shipping_method' => array(
                'label' => 'Shipping Method',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'shipping_cost' => array(
                'label' => 'Shipping Cost',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'tax_amount' => array(
                'label' => 'Tax Amount',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            ),
            'discount_amount' => array(
                'label' => 'Discount Amount',
                'required' => false,
                'always_include' => false,
                'type' => 'order'
            )
        );
        
        return array_merge($checkout_fields, $additional_fields);
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
        foreach ($this->default_fields as $field_key => $field_info) {
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
        
        // Set default selected fields (all default fields)
        if (!get_option('wpmethods_wc_gs_selected_fields')) {
            $default_fields = array_keys($this->default_fields);
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
                
            case 'billing_phone':
                return $order->get_billing_phone();
                
            case 'billing_address':
                $address = sprintf(
                    "%s, %s, %s, %s, %s",
                    $order->get_billing_address_1(),
                    $order->get_billing_address_2(),
                    $order->get_billing_city(),
                    $order->get_billing_state(),
                    $order->get_billing_postcode()
                );
                $address = preg_replace('/,\s*,/', ',', $address);
                return trim($address, ', ');
                
            case 'product_name':
                $product_names = array();
                foreach ($order->get_items() as $item) {
                    $product_names[] = $item->get_name();
                }
                return implode(', ', $product_names);
                
            case 'order_amount':
                return $order->get_total();
                
            case 'order_status':
                return $order->get_status();
                
            case 'order_date':
                return $order->get_date_created()->format('Y-m-d H:i:s');
                
            case 'product_categories':
                return $this->wpmethods_get_order_categories($order);
                
            case 'shipping_class':
                return $this->wpmethods_get_order_shipping_classes($order);
                
            case 'payment_method_title':
                // Try different methods to get payment method title
                $title = $order->get_payment_method_title();
                if (empty($title)) {
                    $payment_method = $order->get_payment_method();
                    if ($payment_method) {
                        // Get payment gateway title
                        $gateways = WC()->payment_gateways()->payment_gateways();
                        if (isset($gateways[$payment_method])) {
                            $title = $gateways[$payment_method]->get_title();
                        }
                    }
                }
                return $title ? $title : '';
                
            case 'payment_method':
                return $order->get_payment_method();
                
            case 'customer_note':
                return $order->get_customer_note();
                
            case 'transaction_id':
                return $order->get_transaction_id();
                
            case 'customer_ip':
                return $order->get_customer_ip_address();
                
            case 'customer_user_agent':
                return $order->get_customer_user_agent();
                
            case 'shipping_method':
                // Get shipping method from order
                $shipping_methods = array();
                foreach ($order->get_shipping_methods() as $shipping_item) {
                    // Try to get method title, then method ID
                    $method_title = $shipping_item->get_method_title();
                    $method_id = $shipping_item->get_method_id();
                    
                    if (!empty($method_title)) {
                        $shipping_methods[] = $method_title;
                    } elseif (!empty($method_id)) {
                        $shipping_methods[] = $method_id;
                    }
                }
                
                // If no shipping methods found, try order's shipping method
                if (empty($shipping_methods)) {
                    $order_shipping_method = $order->get_shipping_method();
                    if (!empty($order_shipping_method)) {
                        $shipping_methods[] = $order_shipping_method;
                    }
                }
                
                return !empty($shipping_methods) ? implode(', ', $shipping_methods) : '';
                
            case 'shipping_cost':
                return $order->get_shipping_total();
                
            case 'tax_amount':
                return $order->get_total_tax();
                
            case 'discount_amount':
                return $order->get_discount_total();
                
            // Handle billing fields
            case strpos($field_key, 'billing_') === 0:
                $method_name = 'get_' . $field_key;
                if (method_exists($order, $method_name)) {
                    return $order->$method_name();
                }
                break;
                
            // Handle shipping fields
            case strpos($field_key, 'shipping_') === 0:
                $method_name = 'get_' . $field_key;
                if (method_exists($order, $method_name)) {
                    return $order->$method_name();
                }
                break;
        }
        
        return null;
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
     * Get shipping classes from order products
     */
    private function wpmethods_get_order_shipping_classes($order) {
        $all_shipping_classes = array();
        
        foreach ($order->get_items() as $item) {
            $product = $item->get_product();
            
            if ($product) {
                $shipping_class = $product->get_shipping_class();
                if ($shipping_class) {
                    // Get shipping class term name
                    $shipping_class_term = get_term_by('slug', $shipping_class, 'product_shipping_class');
                    if ($shipping_class_term && !is_wp_error($shipping_class_term)) {
                        $all_shipping_classes[] = $shipping_class_term->name;
                    } else {
                        $all_shipping_classes[] = $shipping_class;
                    }
                }
            }
        }
        
        $all_shipping_classes = array_unique($all_shipping_classes);
        
        return implode(', ', $all_shipping_classes);
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
        
        // Create admin.js content inline
        $admin_js = '
        jQuery(document).ready(function($) {
            // Generate Google Apps Script
            $("#wpmethods-generate-script").on("click", function(e) {
                e.preventDefault();
                
                var button = $(this);
                var originalText = button.text();
                
                button.text("Generating...").prop("disabled", true);
                
                // Get selected fields
                var selectedFields = [];
                $("input[name=\'wpmethods_wc_gs_selected_fields[]\']:checked").each(function() {
                    selectedFields.push($(this).val());
                });
                
                $.ajax({
                    url: ajaxurl,
                    type: "POST",
                    data: {
                        action: "wpmethods_generate_google_script",
                        nonce: "' . wp_create_nonce('wpmethods_generate_script_nonce') . '",
                        fields: selectedFields
                    },
                    success: function(response) {
                        if (response.success) {
                            $("#wpmethods-generated-script").val(response.data.script);
                            $("#wpmethods-script-output").show();
                            $("html, body").animate({
                                scrollTop: $("#wpmethods-script-output").offset().top - 100
                            }, 500);
                        } else {
                            alert("Error generating script. Please try again.");
                        }
                    },
                    error: function() {
                        alert("Error generating script. Please try again.");
                    },
                    complete: function() {
                        button.text(originalText).prop("disabled", false);
                    }
                });
            });
            
            // Copy to clipboard
            $("#wpmethods-copy-script").on("click", function() {
                var textarea = $("#wpmethods-generated-script")[0];
                textarea.select();
                document.execCommand("copy");
                
                $("#wpmethods-copy-status").show().fadeOut(2000);
            });
            
            // Handle required fields
            $("input[name=\'wpmethods_wc_gs_selected_fields[]\']").each(function() {
                if ($(this).is(":disabled")) {
                    $(this).prop("checked", true);
                }
            });
            
            // Prevent unchecking required fields
            $("input[name=\'wpmethods_wc_gs_selected_fields[]\']").on("change", function() {
                if ($(this).is(":disabled") && !$(this).is(":checked")) {
                    $(this).prop("checked", true);
                }
            });
        });
        ';
        
        // Add inline script
        wp_add_inline_script('jquery', $admin_js);
        
        // Add inline CSS
        $admin_css = '
        .wpmethods-settings-wrapper {
            max-width: 1200px;
        }
        .wpmethods-field-group {
            background: #f9f9f9;
            border: 1px solid #ddd;
            border-radius: 4px;
            padding: 15px;
            margin-bottom: 20px;
        }
        .wpmethods-field-group h3 {
            margin-top: 0;
            padding-bottom: 10px;
            border-bottom: 1px solid #ddd;
        }
        .wpmethods-field-checkboxes {
            max-height: 300px;
            overflow-y: auto;
            padding: 10px;
        }
        .wpmethods-field-item {
            margin-bottom: 8px;
            padding: 5px;
            background: white;
            border-radius: 3px;
        }
        .wpmethods-field-item label {
            display: flex;
            align-items: center;
            cursor: pointer;
        }
        .wpmethods-field-item input[type="checkbox"] {
            margin-right: 8px;
        }
        .wpmethods-field-item .required {
            color: #d63638;
            font-weight: bold;
            margin-left: 5px;
        }
        .wpmethods-generated-script {
            font-family: "Courier New", monospace;
            font-size: 12px;
            line-height: 1.4;
        }
        #wpmethods-script-output {
            animation: fadeIn 0.3s ease-in;
        }
        @keyframes fadeIn {
            from { opacity: 0; }
            to { opacity: 1; }
        }
        ';
        
        wp_add_inline_style('wp-admin', $admin_css);
    }
    
    /**
     * Initialize settings with tabs
     */
    public function wpmethods_settings_init() {
        // Main settings
        register_setting('wpmethods_wc_gs_main_settings', 'wpmethods_wc_gs_order_statuses', array($this, 'wpmethods_sanitize_array'));
        register_setting('wpmethods_wc_gs_main_settings', 'wpmethods_wc_gs_script_url', 'esc_url_raw');
        register_setting('wpmethods_wc_gs_main_settings', 'wpmethods_wc_gs_product_categories', array($this, 'wpmethods_sanitize_array'));
        
        // Checkout fields settings
        register_setting('wpmethods_wc_gs_fields_settings', 'wpmethods_wc_gs_selected_fields', array($this, 'wpmethods_sanitize_array'));
        
        // Main settings section
        add_settings_section(
            'wpmethods_wc_gs_main_section',
            'Main Configuration',
            array($this, 'wpmethods_main_section_callback'),
            'wpmethods_wc_gs_main_settings'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_order_statuses',
            'Trigger Order Statuses',
            array($this, 'wpmethods_order_statuses_render'),
            'wpmethods_wc_gs_main_settings',
            'wpmethods_wc_gs_main_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_product_categories',
            'Product Categories Filter',
            array($this, 'wpmethods_product_categories_render'),
            'wpmethods_wc_gs_main_settings',
            'wpmethods_wc_gs_main_section'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_script_url',
            'Google Apps Script URL',
            array($this, 'wpmethods_script_url_render'),
            'wpmethods_wc_gs_main_settings',
            'wpmethods_wc_gs_main_section'
        );
        
        // Checkout fields section
        add_settings_section(
            'wpmethods_wc_gs_fields_section',
            'Checkout Fields Selection',
            array($this, 'wpmethods_fields_section_callback'),
            'wpmethods_wc_gs_fields_settings'
        );
        
        add_settings_field(
            'wpmethods_wc_gs_selected_fields',
            'Select Fields to Send',
            array($this, 'wpmethods_selected_fields_render'),
            'wpmethods_wc_gs_fields_settings',
            'wpmethods_wc_gs_fields_section'
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
     * Main section callback
     */
    public function wpmethods_main_section_callback() {
        echo '<p>Configure the main settings for Google Sheets integration.</p>';
    }
    
    /**
     * Fields section callback
     */
    public function wpmethods_fields_section_callback() {
        echo '<p>Select which checkout fields should be sent to Google Sheets. Required fields are always included.</p>';
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
        
        foreach ($all_statuses as $status => $label) {
            $checked = in_array($status, $selected_statuses) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 5px;">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_order_statuses[]" 
                           value="<?php echo esc_attr($status); ?>" <?php echo $checked; ?>>
                    <?php echo esc_html($label); ?>
                </label>
            </div>
            <?php
        }
        echo '<p class="description">Select order statuses that should trigger sending data to Google Sheets</p>';
    }
    
    /**
     * Product Categories field render
     */
    public function wpmethods_product_categories_render() {
        $selected_categories = $this->wpmethods_get_selected_categories();
        $all_categories = $this->wpmethods_get_product_categories();
        
        if (empty($all_categories)) {
            echo '<p>No product categories found. Please create some product categories in WooCommerce.</p>';
            return;
        }
        
        echo '<div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin-bottom: 10px;">';
        
        foreach ($all_categories as $cat_id => $cat_name) {
            $checked = in_array($cat_id, $selected_categories) ? 'checked' : '';
            ?>
            <div style="margin-bottom: 5px;">
                <label>
                    <input type="checkbox" name="wpmethods_wc_gs_product_categories[]" 
                           value="<?php echo esc_attr($cat_id); ?>" <?php echo $checked; ?>>
                    <?php echo esc_html($cat_name); ?>
                </label>
            </div>
            <?php
        }
        
        echo '</div>';
        echo '<p class="description">Select product categories. Orders will only be sent if they contain at least one product from selected categories. Leave empty to include all categories.</p>';
    }
    
    /**
     * Selected fields render
     */
    public function wpmethods_selected_fields_render() {
        $selected_fields = $this->wpmethods_get_selected_fields();
        
        // Group fields by type
        $grouped_fields = array(
            'default' => array(),
            'billing' => array(),
            'shipping' => array(),
            'order' => array()
        );
        
        foreach ($this->available_fields as $field_key => $field_info) {
            $type = isset($field_info['type']) ? $field_info['type'] : 'default';
            if (!isset($grouped_fields[$type])) {
                $grouped_fields[$type] = array();
            }
            $grouped_fields[$type][$field_key] = $field_info;
        }
        
        // Display fields by group
        foreach ($grouped_fields as $group_name => $fields) {
            if (empty($fields)) continue;
            
            $group_label = ucfirst($group_name) . ' Fields';
            if ($group_name === 'default') $group_label = 'Default Fields';
            
            echo '<h4>' . esc_html($group_label) . '</h4>';
            echo '<div style="max-height: 200px; overflow-y: auto; border: 1px solid #ddd; padding: 10px; margin-bottom: 20px;">';
            
            foreach ($fields as $field_key => $field_info) {
                $checked = in_array($field_key, $selected_fields) ? 'checked' : '';
                $disabled = (isset($field_info['always_include']) && $field_info['always_include']) ? 'disabled' : '';
                $required = (isset($field_info['required']) && $field_info['required']) ? ' <span style="color:red">*</span>' : '';
                $title = isset($field_info['always_include']) && $field_info['always_include'] ? 'title="This field is required and cannot be disabled"' : '';
                ?>
                <div style="margin-bottom: 5px;">
                    <label <?php echo $title; ?>>
                        <input type="checkbox" name="wpmethods_wc_gs_selected_fields[]" 
                               value="<?php echo esc_attr($field_key); ?>" 
                               <?php echo $checked; ?> <?php echo $disabled; ?>>
                        <?php echo esc_html($field_info['label']) . $required; ?>
                        <?php if ($disabled): ?>
                            <em>(Required)</em>
                        <?php endif; ?>
                    </label>
                </div>
                <?php
            }
            
            echo '</div>';
        }
        
        echo '<p class="description">Select fields to include in Google Sheets. Required fields (*) are always included.</p>';
        ?>
        <div style="margin-top: 20px; padding: 15px; background: #f5f5f5; border-radius: 4px;">
            <h4>Generate Google Apps Script</h4>
            <p>Click the button below to generate a Google Apps Script code based on your selected fields:</p>
            <button type="button" id="wpmethods-generate-script" class="button button-primary">
                Generate Google Apps Script
            </button>
            <div id="wpmethods-script-output" style="margin-top: 15px; display: none;">
                <textarea id="wpmethods-generated-script" style="width: 100%; height: 400px; font-family: monospace;" readonly></textarea>
                <p style="margin-top: 10px;">
                    <button type="button" id="wpmethods-copy-script" class="button button-secondary">
                        Copy to Clipboard
                    </button>
                    <span id="wpmethods-copy-status" style="margin-left: 10px; color: green; display: none;">Copied!</span>
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
        <input type="url" name="wpmethods_wc_gs_script_url" 
               value="<?php echo esc_url($value); ?>" 
               style="width: 500px;" 
               placeholder="https://script.google.com/macros/s/...">
        <p class="description">Enter your Google Apps Script web app URL. Get this from Google Apps Script deployment.</p>
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
        
        // Generate Google Apps Script code
        $script = $this->generate_google_apps_script($selected_fields);
        
        wp_send_json_success(array(
            'script' => $script,
            'fields' => $selected_fields
        ));
    }
    
    /**
     * Generate Google Apps Script code based on selected fields
     */
    private function generate_google_apps_script($selected_fields) {
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
     * Settings page with tabs
     */
    public function wpmethods_settings_page() {
        $active_tab = isset($_GET['tab']) ? sanitize_text_field($_GET['tab']) : 'main';
        ?>
        <div class="wrap">
            <h1>WooCommerce to Google Sheets</h1>
            
            <h2 class="nav-tab-wrapper">
                <a href="?page=wpmethods-wc-to-google-sheets&tab=main" class="nav-tab <?php echo $active_tab == 'main' ? 'nav-tab-active' : ''; ?>">
                    Main Settings
                </a>
                <a href="?page=wpmethods-wc-to-google-sheets&tab=checkout-fields" class="nav-tab <?php echo $active_tab == 'checkout-fields' ? 'nav-tab-active' : ''; ?>">
                    Checkout Fields
                </a>
            </h2>
            
            <?php if ($active_tab == 'main'): ?>
                <form action="options.php" method="post">
                    <?php
                    settings_fields('wpmethods_wc_gs_main_settings');
                    do_settings_sections('wpmethods_wc_gs_main_settings');
                    submit_button();
                    ?>
                </form>
                
                <h3>Current Configuration Summary:</h3>
                <?php
                $selected_statuses = get_option('wpmethods_wc_gs_order_statuses', array());
                $selected_categories = $this->wpmethods_get_selected_categories();
                $all_categories = $this->wpmethods_get_product_categories();
                
                if (!is_array($selected_statuses)) {
                    $selected_statuses = maybe_unserialize($selected_statuses);
                    if (!is_array($selected_statuses)) {
                        $selected_statuses = array('completed', 'processing');
                    }
                }
                
                echo '<p><strong>Trigger Statuses:</strong> ';
                if (!empty($selected_statuses)) {
                    $status_labels = array();
                    foreach ($selected_statuses as $status) {
                        $status_labels[] = ucfirst($status);
                    }
                    echo implode(', ', $status_labels);
                } else {
                    echo 'None selected';
                }
                echo '</p>';
                
                echo '<p><strong>Selected Categories:</strong> ';
                if (!empty($selected_categories)) {
                    $category_names = array();
                    foreach ($selected_categories as $cat_id) {
                        if (isset($all_categories[$cat_id])) {
                            $category_names[] = $all_categories[$cat_id];
                        }
                    }
                    echo implode(', ', $category_names);
                } else {
                    echo 'All categories (no filter)';
                }
                echo '</p>';
                
                $selected_fields = $this->wpmethods_get_selected_fields();
                echo '<p><strong>Selected Fields (' . count($selected_fields) . '):</strong> ';
                $field_names = array();
                foreach ($selected_fields as $field_key) {
                    if (isset($this->available_fields[$field_key])) {
                        $field_names[] = $this->available_fields[$field_key]['label'];
                    }
                }
                echo implode(', ', $field_names);
                echo '</p>';
                ?>
                
            <?php elseif ($active_tab == 'checkout-fields'): ?>
                <form action="options.php" method="post">
                    <?php
                    settings_fields('wpmethods_wc_gs_fields_settings');
                    do_settings_sections('wpmethods_wc_gs_fields_settings');
                    submit_button('Save Fields Selection');
                    ?>
                </form>
            <?php endif; ?>
        </div>
        <?php
    }
}

// Initialize the plugin
new WPMethods_WC_To_Google_Sheets();