<?php

/*
  Plugin Name: Convert Docx2post
  Description: Convert Microsoft Word docx or text files to Wordpress posts, pages or custom post types
  Version: 1.4
  Author: Davor Zeljkovic
  Author URI: https://wpsuit.review
 */

if (is_admin()) {
    include dirname(__FILE__) . '/admin.php';
}


    include dirname(__FILE__) . '/functions.php';

