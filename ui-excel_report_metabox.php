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
function ui_excel_report_metabox($target_post, $args) {
	global $excel_report_maker;

	$target_post_id = $target_post->ID;
	$target_post_type = get_post_meta($target_post_id, 'excel-report-maker-post_type', true);
	$excel_file = get_post_meta($target_post_id, 'excel-report-maker-excel_file_ext', true);
	$worksheet = get_post_meta($target_post_id, 'excel-report-maker-worksheet', true);
	$datas = get_post_meta($target_post_id, 'excel-report-maker-datas', true);

	$show_list_column = get_post_meta($target_post_id, 'excel-report-maker-show_list_column', true);
	$show_in_row_action = get_post_meta($target_post_id, 'excel-report-maker-show_in_row_action', true);
	$show_in_edit = get_post_meta($target_post_id, 'excel-report-maker-show_in_edit', true);


	$array_error_msgs = array(
		0 => esc_html('You must upload Excel Template File', 'excel-report-maker'),
		1 => esc_html('You must input Target Meta Key', 'excel-report-maker'),
		2 => esc_html('You must input Target Cell', 'excel-report-maker'),
	);
	?>
	<script type="text/javascript">
		var array_error_msgs = <?php echo json_encode( $array_error_msgs );?>;
	</script>
	<input type="hidden" name="excel-report-maker-action" value="entry_save" />
	<input type="hidden" name="excel-report-maker-nonce" value="<?php echo $excel_report_maker->get_nonce();?>" />
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Target Posttype', 'excel-report-maker'); ?><span class="red">*</span></label>
		<div class="elm">
			<select name="excel-report-maker-target_posttype">
				<?php
				$target_posttypes = get_option('excel-report-maker-targets');
				foreach($target_posttypes as $post_type => $is_target) {
					if($is_target) {
						$obj_post_type = get_post_type_object( $post_type );
						?>
						<option value="<?php echo esc_html( $post_type );?>" <?php echo ($target_post_type == $post_type)?'selected':'';?> ><?php echo esc_html($obj_post_type->labels->name);?></option>

						<?php
					}
				}
				?>

			</select>
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Excel Template File', 'excel-report-maker'); ?><span class="red">*</span></label>
		<div class="elm">
			<?php
			if($excel_file) {
				?>
				<input type="hidden" name="excel-report-maker-has-file" id="excel-report-maker-has-file" value="yes" />
				<?php esc_html_e('Change', 'excel-report-maker'); ?>: <input type="file" name="excel-report-maker-excel_file" id="excel-report-maker-excel_file" value="" />
				<?php
			} else {
				?>
				<input type="hidden" name="excel-report-maker-has-file" id="excel-report-maker-has-file" value="no" />
				<input type="file" name="excel-report-maker-excel_file" id="excel-report-maker-excel_file" value="" />
				<?php
			}
			?>
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Wordsheet Name', 'excel-report-maker'); ?></label>
		<div class="elm">
			<input type="text" name="excel-report-maker-worksheet" value="<?php echo esc_attr( $worksheet );?>" /><br />
			<p class="desc"><?php esc_html_e('If empty, then use first worksheet.', 'excel-report-maker'); ?></p>
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Add Column And Show in List', 'excel-report-maker'); ?></label>
		<div class="elm">
			<input type="checkbox" name="excel-report-maker-show_list_column" value="1" <?php echo $show_list_column?'checked':'';?> />
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Show in Row action', 'excel-report-maker'); ?></label>
		<div class="elm">
			<input type="checkbox" name="excel-report-maker-show_in_row_action" value="1" <?php echo $show_in_row_action?'checked':'';?> />
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Show in Edit Page', 'excel-report-maker'); ?></label>
		<div class="elm">
			<input type="checkbox" name="excel-report-maker-show_in_edit" value="1" <?php echo $show_in_edit?'checked':'';?> />
		</div>
	</fieldset>
	<fieldset>
		<label style="margin-bottom: 0.5em;"><?php esc_html_e('Data Set To Cell', 'excel-report-maker'); ?></label>
		<div class="elm">
			<table class="tbl-datas">
				<thead>
					<th><?php esc_html_e('Target Meta Key', 'excel-report-maker'); ?><span class="red">*</span></th>
					<th><?php esc_html_e(' to ', 'excel-report-maker'); ?></th>
					<th><?php esc_html_e('Target Cell( ex. A3 )', 'excel-report-maker'); ?><span class="red">*</span></th>
					<th>delete</th>
				</thead>
			<?php
			if($datas) {
				// データ出力
				$datas = maybe_unserialize($datas);
				foreach($datas as $data) {
					?>
					<tr class="">
						<td><input type="text" class="excel-report-maker-metakey" name="excel-report-maker-metakey[]" value="<?php echo esc_attr( $data['meta_key'] );?>" /></td>
						<td><?php esc_html_e(' to ', 'excel-report-maker'); ?></td>
						<td><input type="text" class="excel-report-maker-cell" name="excel-report-maker-cell[]" value="<?php echo esc_attr( $data['target_cell'] );?>" /></td>
						<td><input type="button" class="excel-report-maker-row_delete button" value="<?php esc_html_e('Delete Row', 'excel-report-maker'); ?>" /></td>
					</tr>
					<?php
				}
			}

			// 空の1行
			?>
				<tr class="excel-report-maker-template">
					<td><input type="text" class="excel-report-maker-metakey" name="excel-report-maker-metakey[]" value="" /></td>
					<td><?php esc_html_e(' to ', 'excel-report-maker'); ?></td>
					<td><input type="text" class="excel-report-maker-cell" name="excel-report-maker-cell[]" value="" /></td>
					<td><input type="button" class="excel-report-maker-row_delete button" value="<?php esc_html_e('Delete Row', 'excel-report-maker'); ?>" /></td>
				</tr>

			<?php


			?>
			</table>
			<p>
				<input type="button" class="excel-report-maker-row-add button button-primary" value="<?php esc_html_e('Add Row', 'excel-report-maker'); ?>" />
			</p>
		</div>
	</fieldset>


	<?php
}
