<?php
if ( ! defined( 'ABSPATH' ) ) {
    exit; // Exit if accessed directly
}
spl_autoload_register(function($class) {
    $prefix = 'UGSIW\\';
    $base_dir = UGSIW_WPMETHODS_PATH . 'includes/';
    $len = strlen($prefix);
    if (strncmp($prefix, $class, $len) !== 0) {
        return;
    }

    $relative_class = substr($class, $len);
    $file = $base_dir . 'class-' . strtolower(str_replace('_', '-', $relative_class)) . '.php';

    if (file_exists($file)) {
        require $file;
    }
});