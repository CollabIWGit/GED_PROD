import * as $ from 'jquery';
import styles from './DocDetailsWebPart.module.scss';

$(function () {
    // $("#panel").hide();

    // $("#btn_filters").on('click',function () {
    //     var visible = $("#panel").is(":hidden");
    //     if (visible)
    //         $("#btn_filters").html('<i class="fa fa-times fa-fw"></i>&nbsp;&nbsp;Filters');
    //     else
    //         $("#btn_filters").html('<i class="fa fa-plus fa-fw"></i>&nbsp;&nbsp;Filters');
    //     $("#panel").slideToggle("fast");
    // });

    // $("#tbl_contract_mgt").DataTable({
    //     initComplete: function () {
    //         //add document date drop down
    //         $('<label>&nbsp;&nbsp;Country:&nbsp;</label>').appendTo('#tbl_contract_mgt_filter');
    //         this.api().columns([5]).every(function () {
    //             var column = this;
    //             var select = $('<select class="docdate_filter"><option value="">Show All</option></select>')
    //                 .appendTo('#tbl_contract_mgt_filter')
    //                 .on('change', function () {
    //                     var val = $.fn.dataTable.util.escapeRegex(
    //                         $(this).val()
    //                     );
    //                     column
    //                         .search(val ? '^' + val + '$' : '', true, false)
    //                         .draw();
    //                 });

    //             column.data().unique().sort().each(function (d, j) {
    //                 select.append('<option value="' + d + '">' + d + '</option>')
    //             });
    //             //avoid select click to mess with the ordering of the table
    //             $(select).on('click', function (e) {
    //                 e.stopPropagation();
    //             });
    //         });
    //         // this.api().column(2).every(function () {
    //         //     var column = this;
    //         //     $("#drp_supplier")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_suppliers").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //         // this.api().column(6).every(function () {
    //         //     var column = this;
    //         //     $("#drp_contract_owner")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_owners").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //         // this.api().column(9).every(function () {
    //         //     var column = this;
    //         //     $("#drp_start_date_month")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_start_date_months").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //         // this.api().column(10).every(function () {
    //         //     var column = this;
    //         //     $("#drp_end_date_month")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_end_date_months").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //         // this.api().column(11).every(function () {
    //         //     var column = this;
    //         //     $("#drp_start_date_year")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_start_date_years").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //         // this.api().column(12).every(function () {
    //         //     var column = this;
    //         //     $("#drp_end_date_year")
    //         //         .on('keyup', function () {
    //         //             column
    //         //                 .search($(this).val())
    //         //                 .draw();
    //         //         });

    //         //     column.data().unique().sort().each(function (d, j) {
    //         //         $("#select_end_date_years").append(`<option value="${d}">${d}</option>`)
    //         //     });
    //         // });
    //     },
    //     paging: true,
    //     info: true,
    //     searching: true,
    //     search: {
    //         smart: false,
    //         regex: true
    //     },
    //     responsive: true,
    //     columnDefs: [
    //         {
    //             "targets": [0, 7, 8, 10, 11, 12, 13, 14],
    //             "visible": false
    //         },
    //         { orderable: false, targets: [15,16] }
    //     ],
    //     order: [[0, "desc"]]
    // });

    function floatLabel() {
        $('.floatLabel').each(function () {
            var $this = $(this);
            // on focus add cladd active to label
            $this.focus(function () {
                $this.next().addClass(`${styles.active}`);
            });
            //on blur check field and remove class if needed
            $this.blur(function () {
                if ($this.val() === '' || $this.val() === 'blank') {
                    $this.next().removeClass();
                }
            });
        });
    }
    // just add a class of "floatLabel to the input field!"
    floatLabel();

    function floatLabel2() {
        $('.floatLabel2').each(function () {
            var $this = $(this);
            // on focus add cladd active to label
            $this.focus(function () {
                $this.next().next().addClass(`${styles.active}`);
            });
            //on blur check field and remove class if needed
            $this.blur(function () {
                if ($this.val() === '' || $this.val() === 'blank') {
                    $this.next().next().removeClass();
                }
            });
        });
    }
    // just add a class of "floatLabel2 to the input field!"
    floatLabel2();
});