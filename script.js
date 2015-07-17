(function($) {
	$(function() {
		if($('#excel_report_metabox_excel_report').size() != 0) {
			// metabox
			$('.excel-report-maker-row-add').on("click", function() {
				// add row
				$('tr.excel-report-maker-template').clone(true).appendTo($('table.tbl-datas')).removeClass('excel-report-maker-template');

				return false;
			});


			$('.excel-report-maker-row_delete').on("click", function(){
				var $parent_tr = $(this).parents('tr');
				$parent_tr.remove();

				return false;
			});

			$('#post').attr('enctype','multipart/form-data');

			$('#post').on("submit", function() {
				// file
				if($('#excel-report-maker-has-file').val() == 'no') {
					// ng
					if($('#excel-report-maker-excel_file').val() == '') {
						$('#excel-report-maker-excel_file').addClass('error');
						alert(array_error_msgs[0]);
						return false;
					}
				}
				$('#excel-report-maker-excel_file').removeClass('error');

				var has_error = false;
				$('.excel-report-maker-metakey').each(function() {
					if($(this).parents('tr').hasClass('excel-report-maker-template')) {

					} else {
						if($(this).val() == '') {
							$(this).addClass('error');
							has_error = true;
						} else {
							$(this).removeClass('error');
						}
					}
				});
				if(has_error) {
					alert(array_error_msgs[1]);
					has_error = true;
					return false;
				}

				has_error = false;
				$('.excel-report-maker-cell').each(function() {
					if($(this).parents('tr').hasClass('excel-report-maker-template')) {

					} else {
						if ($(this).val() == '') {
							$(this).addClass('error');
							has_error = true;
						} else {
							$(this).removeClass('error');
						}
					}
				});
				if(has_error) {
					alert(array_error_msgs[2]);
					has_error = true;
					return false;
				}

				return true;
			});


		}





	});


})(jQuery);