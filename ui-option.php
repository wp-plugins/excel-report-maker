<?php
/**
 *
 *
 * Created by PhpStorm.
 * Author: Eyeta Co.,Ltd.(http://www.eyeta.jp)
 *
 */

namespace excel_report_maker;

/**
 * setting page
 */
function ui_option() {
	global $excel_report_maker;

	// save
	if('save' == $_REQUEST['excel-report-maker-action']) {
		// save
		if(!$excel_report_maker->verify_nonce($_REQUEST['excel-report-maker-nonce'])) {
			// nonce NG
			wp_die(__('Maybe Timeout', 'excel-report-maker'));
		}

		// 保存
		if(isset($_REQUEST['target_posttype'])) {
			$target_posttypes = array();
			$post_types = wp_list_filter( get_post_types() );
			if(is_array($_REQUEST['target_posttype'])) {
				foreach($_REQUEST['target_posttype'] as $key => $post_type) {
					if(in_array($post_type, $post_types)) {
						$target_posttypes[$post_type] = true;
					}
				}

			} else {
				if(in_array($_REQUEST['target_posttype'], $post_types)) {
					$target_posttypes[$_REQUEST['target_posttype']] = true;
				}
			}
			update_option('excel-report-maker-targets', $target_posttypes);
		} else {
			// 対象なし
			update_option('excel-report-maker-targets', array());
		}

		//
		if(isset($_REQUEST['use-acf']) && 'use-acf' == $_REQUEST['use-acf']) {
			update_option( 'excel-report-maker-use-acf', 1 );
		} else {
			delete_option( 'excel-report-maker-use-acf' );
		}
	} else {
		$target_posttypes = get_option('excel-report-maker-targets');
	}
	if(!is_array($target_posttypes)) {
		$target_posttypes = array();
	}
	$use_acf = get_option( 'excel-report-maker-use-acf' );

	?>
	<div class="wrap">

		<h2><?php esc_html_e('Excel Reports Setting', 'excel-report-maker'); ?></h2>

		<div id="tabs">
			<!--<ul>
				<li><a href="#excel-report-maker-opt1"><?php esc_html_e("Global Setting", 'excel-report-maker'); ?></a></li>
				<li><a href="#fragment-3">Three</a></li>
			</ul>-->
			<div id="excel-report-maker-opt1">
				<form action="<?php echo admin_url('options-general.php?page=excel-report-maker');?>" method="post" name="frm-excel-report-maker-option">
					<h3><?php esc_html_e("Target Posttype", 'excel-report-maker'); ?></h3>
					<input type="hidden" name="excel-report-maker-nonce" value="<?php echo $excel_report_maker->get_nonce();?>" />
					<input type="hidden" name="excel-report-maker-action" value="save" />
					<?php

					$post_types = wp_list_filter( get_post_types() );
					foreach( $post_types as $post_type) {
						$obj_post_type = get_post_type_object( $post_type );
						?>
						<label for="target_posttype_<?php echo esc_attr( $post_type );?>"><input type="checkbox" name="target_posttype[]" id="target_posttype_<?php echo esc_attr( $post_type );?>" value="<?php echo esc_html( $post_type );?>" <?php echo ($excel_report_maker->is_target_posttype( $post_type ))?'checked':'';?>/> <?php echo esc_html($obj_post_type->labels->name);?></label><br />
					<?php

					}
						?>

					<h3><?php esc_html_e("Use ACF", 'excel-report-maker'); ?></h3>
					<label for="use-acf"><input type="checkbox" name="use-acf" id="use-acf" value="use-acf" <?php echo ($use_acf)?'checked':'';?>/> <?php esc_html_e("Use ACF", 'excel-report-maker'); ?></label><br />



					<br />
					<br />
					<input type="submit" name="btn-submit" class="button button-primary" value="<?php esc_html_e('Save', 'excel-report-maker'); ?>" />


				</form>
			</div>
			<!--<div id="fragment-3">
				Lorem ipsum dolor sit amet, consectetuer adipiscing elit, sed diam nonummy nibh euismod tincidunt ut laoreet dolore magna aliquam erat volutpat.
			</div>-->
		</div>

	</div>



<?php
}
