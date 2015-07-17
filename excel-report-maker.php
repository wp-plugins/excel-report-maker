<?php
/*
Plugin Name: Excel-report-maker
Version: 0.1.2
Description: set post and postmeta values to excel file
Author: Eyeta Co.,Ltd.
Author URI: http://www.eyeta.jp
Plugin URI:
Text Domain: excel-report-maker
Domain Path: /languages
*/

namespace excel_report_maker;
/**
 *
 *
 * Created by PhpStorm.
 * Author: Eyeta Co.,Ltd.(http://www.eyeta.jp)
 *
 */

\register_activation_hook( __FILE__, '\excel_report_maker\excel_report_maker_activate' );
function excel_report_maker_activate() {
	// プラグイン初期化

}

\register_deactivation_hook( __FILE__, '\excel_report_maker\excel_report_maker_deactivate' );
function excel_report_maker_deactivate() {
	// プラグイン削除
}

class excel_report_maker {

	protected $_plugin_dirname;
	protected $_plugin_url;
	protected $_plugin_path;


	public function __construct() {
		// 初期パス等セット
		$this->init();

		// フックセット等
		\add_action('init', array(&$this, 'init_action'));
		\add_action('wp_print_scripts', array(&$this, 'wp_print_scripts'));

		// 管理画面追加
		\add_action('admin_menu', array(&$this, 'admin_menu'));

		// css, js
		\add_action('admin_print_styles', array(&$this, 'head_css'));
		\add_action('admin_print_scripts', array(&$this, 'head_js'));


		// row action
		add_filter( 'post_row_actions', array(&$this, 'post_row_actions'), 10, 2 );

		// column
		add_filter( 'manage_posts_custom_column', array(&$this, 'manage_posts_custom_column'), 10, 2 );
		add_action( 'manage_posts_columns', array(&$this, 'manage_posts_columns'), 10, 2 );

		// save_post
		add_action( 'save_post', array(&$this, 'save_post'), 5, 3);

		// ajax
		add_action('wp_ajax_excel_report_create', array(&$this, 'ajax_excel_report_create_api'));
		// nopriv
		add_action('wp_ajax_nopriv_excel_report_create', array(&$this, 'ajax_nopriv_excel_report_create_api'));

	}

	/**
	 * ajax レポートExcel作成ダウンロード
	 */
	function ajax_excel_report_create_api() {
		if(apply_filters('excel-report-current_user_can', true, $_REQUEST['target_report_id'], $_REQUEST['target_post_id'])) {
			$this->excel_report_create( $_REQUEST['target_report_id'], $_REQUEST['target_post_id'] );
		}
	}
	function ajax_nopriv_excel_report_create_api() {
		if(apply_filters('excel-report-current_user_can', false, $_REQUEST['target_report_id'], $_REQUEST['target_post_id'])) {
			$this->excel_report_create( $_REQUEST['target_report_id'], $_REQUEST['target_post_id'] );
		}
	}

	/**
	 * レポートExcel作成ダウンロード
	 */
	function excel_report_create($report_post_id, $target_post_id) {
		require_once 'phpexcel/Classes/PHPExcel.php';
		require_once 'phpexcel/Classes/PHPExcel/IOFactory.php';

		// 対象レポート
		$report_post_id = intval( $report_post_id );
		$report_post    = \get_post( $report_post_id );
		if ( 'excel_report' != $report_post->post_type ) {
			wp_die( __("Isn't excel report", 'excel-report-maker') . ': ' . $report_post_id );
		}

		// 対象post
		$target_post_id = intval( $target_post_id );
		$target_post     = \get_post( $target_post_id );
		if ( !$this->is_target_posttype($target_post->post_type)) {
			wp_die( __("Isn't posttype for excel report", 'excel-report-maker') . ': ' . $target_post_id );
		}

		// 対象のExcelファイル
		$filename = $this->get_excel_template_filename( $report_post_id );
		$ext = get_post_meta( $report_post_id, 'excel-report-maker-excel_file_ext', true);

		$excel = new \PHPExcel();
		if ( 'xls' == $ext ) {
			$reader = \PHPExcel_IOFactory::createReader( 'Excel5' );
		} else {
			$reader = \PHPExcel_IOFactory::createReader( 'Excel2007' );
		}
		$book = $reader->load( $this->get_upload_path() . $filename );

		//sheet
		$sheet_name = get_post_meta( $report_post_id, 'excel-report-maker-worksheet', true);
		if ( '' == $sheet_name ) {
			$book->setActiveSheetIndex( 0 );
			$sheet = $book->getActiveSheet();
		} else {
			$sheet = $book->getSheetByName( $sheet_name );
		}
		if ( null == $sheet ) {
			wp_die( __("Cant find target worksheet: ", 'excel-report-maker') . ': ' . $sheet_name );
			exit;
		}

		$is_acf = get_option( 'excel-report-maker-use-acf' );
		$is_acf = intval($is_acf);

		// 埋め込み項目を取得
		$array_input_datas = get_post_meta($report_post_id, 'excel-report-maker-datas', true);
		$array_input_datas = maybe_unserialize($array_input_datas);
		foreach ( $array_input_datas as $array_input_data ) {
			// ACF?
			if($is_acf) {
				// ACF get_field
				// acf type
				$obj_field = get_field_object($array_input_data['meta_key'], $target_post_id);
				if($obj_field) {
					switch($obj_field['type']) {
						default:
							$val = get_field($array_input_data['meta_key'], $target_post_id);
							$sheet->setCellValue( $array_input_data['target_cell'], $val );
							break;
					}
				}

			} else {
				// get_post_meta
				$val = get_post_meta($target_post_id, $array_input_data['meta_key'], true);

				$sheet->setCellValue( $array_input_data['target_cell'], $val );
			}
		}

		// その他データ・セット
		$sheet = apply_filters('excel-report-before_save', $sheet, $report_post_id, $target_post_id );

		// 保存
		//$this->get_upload_path() . $filename
		$filename = $report_post_id . '_' . $target_post_id . '.' . $ext;
		if ( 'xls' == $ext ) {
			$writer = \PHPExcel_IOFactory::createWriter( $book, 'Excel5' );
		} else {
			$writer = \PHPExcel_IOFactory::createWriter( $book, 'Excel2007' );
		}
		$writer->save( $this->get_upload_path() . $filename );
		header( "Content-Type: application/octet-stream" );//ダウンロードの指示
		header( 'Content-Disposition: attachment; filename*=UTF-8\'\'' . rawurlencode( get_the_title( $report_post_id ) . '_' . get_the_title( $target_post_id ) . '.' . $ext ) );
		header( "Content-Length: " . filesize( $this->get_upload_path() . $filename ) );//ダウンロードするファイルのサイズ
		ob_end_clean();//ファイル破損エラー防止
		readfile( $this->get_upload_path() . $filename );

		unlink($this->get_upload_path() . $filename);

		exit;
	}


	/**
	 * 投稿保存時にexcel_reportだったらメタ情報を保存
	 *
	 *
	 * @param $post_id
	 *
	 */
	function save_post( $post_id, $target_post, $update ) {
		global $excel_report_maker;

		if($target_post->post_type == 'excel_report') {
			if(!isset($_REQUEST['excel-report-maker-action']) || 'entry_save' != $_REQUEST['excel-report-maker-action']) {
				return $post_id;
			}

			if(!isset($_REQUEST['excel-report-maker-nonce']) || false == $excel_report_maker->verify_nonce($_REQUEST['excel-report-maker-nonce'])) {
				// nonce NG
				return $post_id;
			}

			$target_posttype = isset($_REQUEST['excel-report-maker-target_posttype'])?$_REQUEST['excel-report-maker-target_posttype']:'';
			if(!$this->is_target_posttype($target_posttype)) {
				return $post_id;
			}
			update_post_meta($post_id, 'excel-report-maker-post_type', $target_posttype);

			$target_worksheet = isset($_REQUEST['excel-report-maker-worksheet'])?$_REQUEST['excel-report-maker-worksheet']:'';
			update_post_meta($post_id, 'excel-report-maker-worksheet', $target_worksheet);

			// excel-report-maker-show_list_column
			$show_list_column = isset($_REQUEST['excel-report-maker-show_list_column'])?$_REQUEST['excel-report-maker-show_list_column']:'';
			update_post_meta($post_id, 'excel-report-maker-show_list_column', $show_list_column);

			$show_in_row_action = isset($_REQUEST['excel-report-maker-show_in_row_action'])?$_REQUEST['excel-report-maker-show_in_row_action']:'';
			update_post_meta($post_id, 'excel-report-maker-show_in_row_action', $show_in_row_action);

			$show_in_edit = isset($_REQUEST['excel-report-maker-show_in_edit'])?$_REQUEST['excel-report-maker-show_in_edit']:'';
			update_post_meta($post_id, 'excel-report-maker-show_in_edit', $show_in_edit);

			if(isset($_REQUEST['excel-report-maker-metakey']) && is_array($_REQUEST['excel-report-maker-metakey'])
				&& isset($_REQUEST['excel-report-maker-cell']) && is_array($_REQUEST['excel-report-maker-cell'])) {
				// data保存 excel-report-maker-datas
				$array_datas = array();
				foreach($_REQUEST['excel-report-maker-metakey'] as $key => $metakey) {
					if(isset($_REQUEST['excel-report-maker-cell'][$key]) && '' != $_REQUEST['excel-report-maker-cell'][$key]) {
						$array_datas[] = array(
							'meta_key' => $metakey,
							'target_cell' => $_REQUEST['excel-report-maker-cell'][$key]
						);
					}
				}
			}
			//error_log(print_r($array_datas, true));
			update_post_meta($post_id, 'excel-report-maker-datas', serialize($array_datas));

			// file
			//error_log(print_r($_FILES, true));
			if(isset($_FILES['excel-report-maker-excel_file'])) {
				$array_file = $_FILES['excel-report-maker-excel_file'];
				if($array_file['name']) {
					$array_mime = \wp_check_filetype($array_file['name']);

					$upload_path = $this->get_upload_path();

					$filename = $this->get_excel_template_filename($post_id);

					if (!move_uploaded_file($array_file['tmp_name'], $upload_path . $filename)) {
						wp_die(__('Error: file upload failer : ', 'excel-report-maker'));
						die;
					}

					update_post_meta($post_id, 'excel-report-maker-excel_file_ext', $array_mime['ext']);
				}
			}
		}
		return $post_id;
	}

	/**
	 * Excelテンプレートファイル名
	 *
	 * @param $post_id
	 * @param $ext
	 *
	 * @return string
	 */
	function get_excel_template_filename($post_id) {

		return 'template_' . $post_id;

	}


	/**
	 * ファイルアップロード要パス
	 *
	 * @return string
	 */
	function get_upload_path() {

		$uploads = \wp_upload_dir('');
		$upload_path = $uploads['basedir'] . '/excel_report/';
		if(!file_exists($upload_path)) {
			if(!mkdir($upload_path)) {
				wp_die(__('Error: mkdir failer : ', 'excel-report-maker') . $upload_path);
			}
			if(!file_exists($upload_path . '.htaccess')) {
				if(!copy($this->get_plugin_path() . '/htaccess',$upload_path . '.htaccess')) {
					wp_die(__('Error: create htaccess failer : ', 'excel-report-maker') . $upload_path . '.htaccess');
				}
			}
		}

		return $upload_path;

	}


	/**
	 * 管理画面一覧カスタマイズ
	 *
	 * @param $column_name
	 * @param $post_id
	 */
	function manage_posts_custom_column($column_name, $post_id) {
		$target_post = get_post($post_id);
		// 対象のpost_typeが登録されているかどうか
		$report_posts = new \WP_Query(array(
			'post_type' => 'excel_report',
			'post_status' => 'publish',
			'nopaging' => true,
			'meta_query' => array(
				array(
					'key' => 'excel-report-maker-post_type',
					'value' => $target_post->post_type
				),
				array(
					'key' => 'excel-report-maker-show_list_column',
					'value' => 1,
					'type' => 'NUMERIC'
				)
			)
		));
		if($report_posts->have_posts()) {
			$report_posts->the_post();
			echo '<a href="' . $this->get_create_url(get_the_ID(), $post_id ) . '" class="button button-primary">' . esc_html(get_the_title()) . '</a> ';
		}
		wp_reset_postdata();


	}

	function manage_posts_columns($posts_columns, $post_type) {

		// 対象のpost_typeが登録されているかどうか
		$report_posts = new \WP_Query(array(
			'post_type' => 'excel_report',
			'post_status' => 'publish',
			'meta_query' => array(
				array(
					'key' => 'excel-report-maker-post_type',
					'value' => $post_type
				),
				array(
					'key' => 'excel-report-maker-show_list_column',
					'value' => 1,
					'type' => 'NUMERIC'
				)
			)
		));
		if($report_posts->have_posts()) {
			$posts_columns['excel_report'] = __('Excel Reports', 'excel-report-maker');
		}
		wp_reset_postdata();

		return $posts_columns;
	}



	/**
	 * row action add
	 *
	 * @param $actions
	 * @param $post
	 */
	function post_row_actions($actions, $post) {

		// 対象のpost_typeが登録されているかどうか
		$report_posts = new \WP_Query(array(
			'post_type' => 'excel_report',
			'post_status' => 'publish',
			'nopaging' => true,
			'meta_query' => array(
				array(
					'key' => 'excel-report-maker-post_type',
					'value' => $post->post_type
				),
				array(
					'key' => 'excel-report-maker-show_in_row_action',
					'value' => 1,
					'type' => 'NUMERIC'
				)
			)
		));
		while($report_posts->have_posts()) {
			$report_posts->the_post();

			$actions['excel_report_' . get_the_ID()] =
				'<a href="' . $this->get_create_url(get_the_ID(), $post->ID ) . '">' . esc_html(get_the_title()) . '</a>';
		}
		wp_reset_postdata();

		return $actions;

	}

	/*
	 * 管理画面追加
	 */
	function admin_menu () {

		require_once 'ui-option.php';
		//require_once 'ui-metabox.php';

		add_submenu_page( 'options-general.php', __('Excel Reports Setting', 'excel-report-maker'), __('Excel Reports', 'excel-report-maker'), 10, 'excel-report-maker', '\excel_report_maker\ui_option');

		// 投稿metabox
		$post_types = wp_list_filter(get_post_types());
		$target_posttypes = get_option('excel-report-maker-targets');
		if(!is_array($target_posttypes)) {
			$target_posttypes = array();
		}
		foreach( $post_types as $post_type) {
			if('excel_report' != $post_type) {

				if($this->is_target_posttype( $post_type )) {
					add_meta_box(
						'excel_report_metabox_' . $post_type,
						__('Excel Reports', 'excel-report-maker'),
						array(&$this, 'add_meta_box'),
						$post_type,
						'normal',
						'high',
						array('post_type' => $post_type)
					);
				}
			} else {
				// excel_report
				add_meta_box(
					'excel_report_metabox_' . $post_type,
					__('Excel Reports', 'excel-report-maker'),
					array(&$this, 'add_meta_box_excel_report'),
					$post_type,
					'normal',
					'high',
					array('post_type' => $post_type)
				);
			}
		}

	}


	function add_meta_box_excel_report($target_post, $args) {

		require_once "ui-excel_report_metabox.php";

		ui_excel_report_metabox($target_post, $args);

	}


	/**
	 * 印刷ボタンメタボックス追加
	 *
	 * @param $target_post
	 * @param $args
	 */
	function add_meta_box($target_post, $args) {
		if($target_post->post_status != 'publish') {
			return ;
		}

		$target_post_id = $target_post->ID;

		$report_posts = new \WP_Query(array(
			'post_type' => 'excel_report',
			'post_status' => 'publish',
			'nopaging' => true,
			'meta_query' => array(
				array(
					'key' => 'excel-report-maker-post_type',
					'value' => $args['args']['post_type']
				),
				array(
					'key' => 'excel-report-maker-show_in_edit',
					'value' => 1,
					'type' => 'NUMERIC'
				)
			)
		));
		while($report_posts->have_posts()) {
			$report_posts->the_post();

				echo '<a href="' . $this->get_create_url(get_the_ID(), $target_post_id ) . '" class="button button-primary">' . esc_html(get_the_title()) . '</a>';
		}
		wp_reset_postdata();

	}

	/**
	 * レポートURLの生成
	 *
	 * @param $report_post_id
	 * @param $target_post_id
	 */
	function get_create_url($report_post_id, $target_post_id) {

		return admin_url( 'admin-ajax.php' ) . '?action=excel_report_create&target_report_id=' . $report_post_id . '&target_post_id=' . $target_post_id . '&dummy=' . date('YmdHis');
	}

	/**
	 * 対象のpost_typeがプラグインの対象と設定されているかをチェック　
	 *
	 * @param $post_type
	 *
	 * @return bool
	 */
	function is_target_posttype($post_type) {
		$target_posttypes = get_option('excel-report-maker-targets');
		if(!is_array($target_posttypes)) {
			$target_posttypes = array();
		}

		return (isset($target_posttypes[$post_type]) && $target_posttypes[$post_type]);

	}


	/**
	 * 管理画面CSS追加
	 */
	function head_css () {
		if(is_admin()) {
			wp_enqueue_style('excel-report-maker', $this->get_plugin_url() . '/style.css');
		}

	}

	/*
	 * 管理画面JS追加
	 */
	function head_js () {
		if( is_admin()) {
			wp_enqueue_script('jquery');
			wp_enqueue_script('excel-report-maker', $this->get_plugin_url() . '/script.js', array('jquery'));

		}
	}

	/*
	 * initフック
	 */
	function init_action() {
		// 各種フックセット
		$labels = array(
			'name' => __('Excel Reports', 'excel-report-maker'),
			'singular_name' => __('Excel Report', 'excel-report-maker'),
			'menu_name' => __('Excel Reports', 'excel-report-maker'),
			'all_items' => __('All Excel Reports', 'excel-report-maker'),
			'add_new' => __('Add New', 'excel-report-maker'),
			'add_new_item' => __('Add Excel Report', 'excel-report-maker'),
			'edit' => __('Edit', 'excel-report-maker'),
			'edit_item' => __('Edit Excel Report', 'excel-report-maker'),
			'new_item' => __('New Excel Report', 'excel-report-maker'),
			'view' => __('View', 'excel-report-maker'),
			'view_item' => __('View Excel Report', 'excel-report-maker'),
			'search_items' => __('Search Excel Report', 'excel-report-maker'),
			'not_found' => __('No Excel Reports found', 'excel-report-maker'),
			'not_found_in_trash' => __('No Excel Reports found in Trash', 'excel-report-maker'),
			'parent' => __('Parent Excel Report', 'excel-report-maker'),
			);

		$args = array(
			'labels' => $labels,
			'description' => '',
			'public' => false,
			'show_ui' => true,
			'has_archive' => false,
			'show_in_menu' => true,
			'exclude_from_search' => true,
			'capability_type' => 'post',
			'map_meta_cap' => true,
			'hierarchical' => false,
			'rewrite' => false,
			'query_var' => false,

			'supports' => array( 'title'/*, 'custom-fields'*/ ),
		);
		register_post_type( 'excel_report', $args );


	}

	/**
	 * javascriptへ変数類をセット
	 */
	function wp_print_scripts() {
		echo '<script type="text/javascript">
			var excel_report_maker_plugin_url = "' . $this->get_plugin_url() . '";
			var excel_report_maker_nonce = "' . $this->get_nonce() . '";
		</script>';
	}

	/*
	 * 初期化
	**/
	function init() {

		$array_tmp = explode(DIRECTORY_SEPARATOR, dirname(__FILE__));
		$this->_plugin_dirname = $array_tmp[count($array_tmp)-1];
		$this->_plugin_url = '/'. PLUGINDIR . '/' . $this->_plugin_dirname;
		$this->_plugin_path = dirname(__FILE__);

		load_plugin_textdomain( 'excel-report-maker', false, basename( dirname( __FILE__ ) ) . '/languages' );

	}
	function get_plugin_url() {
		return $this->_plugin_url;
	}

	function get_plugin_dirname() {
		return $this->_plugin_dirname;
	}

	function get_plugin_path() {
		return $this->_plugin_path;
	}

	/**
	 * nonce取得
	 *
	 * @return string
	 */
	function get_nonce() {
		return \wp_create_nonce(\plugin_basename(__FILE__));
	}

	function verify_nonce($nonce) {
		return wp_verify_nonce($nonce, \plugin_basename(__FILE__));
	}

}

/*
 * エラー関数
 */
function excel_report_maker_log($msg, $level = 'DEBUG') {

	$level_array = array(
		'DEBUG' => 0,
		'DETAIL' => 1,
		'INFO' => 2,
		'ERROR' => 3
	);



	if($level_array[apply_filters('excel_report_maker_log_level', 'DEBUG')] <= $level_array[$level]) {
		if(mb_strlen($msg)< 800) {
			error_log($_SERVER['SERVER_NAME'] . ' : ' . $level . ' : ' . $msg);
		} else {
			$size = mb_strlen($msg);
			for($i=0; $i < $size; $i+=800) {
				error_log($_SERVER['SERVER_NAME'] . ' : ' . $level . ' : ' . mb_substr($msg, $i, 800));
			}
		}
	}


	$error_email = apply_filters('excel_report_maker_error_mail_to', false);
	if($level == 'ERROR' && $error_email !== false) {
		wp_mail( $error_email, 'excel_report_maker ：' . $_SERVER['SERVER_NAME'], $_SERVER['SERVER_NAME'] . ' : ' . $level . ' : ' . $msg, 'From: ' . $error_email );
	}

}


global $excel_report_maker;
$excel_report_maker = new excel_report_maker();
