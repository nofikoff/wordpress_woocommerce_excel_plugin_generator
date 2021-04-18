<?php
/**
 *
 * Plugin Name: NOVIKOV модифицированный : Simple Excel Price list for WooCommerce
 * Plugin URI: Google.com
 * Description: This plugin helps to create a price list of all products available in stock in excel format and allows users to download the file.
 * Text Domain: excel_price
 * License: GPLv2 or later
 * Author: Novikov
 * Version: 9.9999
 */

/** Load composer */
$composer = __DIR__ . '/vendor/autoload.php';
if ( file_exists($composer) ) {
    require_once $composer;
}

if ( is_admin() ) {
    new sepwSettingsPage();
} else {
    new sepwWorker();
}
