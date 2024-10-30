<?php

function docx2post_menu() {
    add_menu_page('convert_docx2post', 'Convert Docx', 'publish_posts', 'convert-docx2post/options.php', '', 'dashicons-media-document', 6);
}
add_action('admin_menu', 'docx2post_menu');
