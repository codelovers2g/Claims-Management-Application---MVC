var checkExportsSync = true;
var ExportDataCount = 0;
var ExportDataArray = new Array();
var ExportDataTempArray = new Array();
var TemporaryLenght;
var OFFSET = 1;
var TotalCharges = 0;
var AccountExports = {
    VendorName: null,
    oTable: null,
    iTableCounter: 1,
    checkCaseId: null,
    recordCount: 0,
    chargesCount: 0,
    innerHtml: function (table_id) {
        return '  <div><table id="tblAccountExports' + table_id + '" class="table table-bordered table-striped table-hover js-basic-example dataTable w-75 mx-auto" >\
                                <thead>\
                                    <tr>\
                                            <th>\CaseChargeId</th >\
                                            <th>\Activity Name</th>\
                                            <th>\Activity Charges</th>\
                                    </tr>\
                                </thead>\
                                <tbody></tbody>\
                            </table></div>';
    },
    init: function () {
        $('#accounts').click();

        // Token to Access QuickBooks
        AccountExports.GetAccessToken();

        // Last Sync Date
        AccountExports.GetLastSyncDate();

        $('#accountExportAnchor').css('color', '#00bcd4');

        $("#btnSearchExport").click(function () {
            var from = $("#exportFromDate").val();
            var to = $("#exportToDate").val();
            if (Date.parse(from) > Date.parse(to)) {
                swal.fire({
                    title: 'Oops..!',
                    text: 'Invalid Date Range.',
                    icon: 'warning',
                    timer: 3000,
                    showConfirmButton: false
                });
            }
            else {

                AccountExports.GetAccountExports(OFFSET);
            }
        });

        // Connect To QuickBooks
        $(document).on('click', '#btnConnectAccountExportsQB', function () {
            AccountExports.ConnectToQuickBook();
        });

        // Sync AccountExports with QuickBooks
        $(document).on('click', '#btnSyncAccountExportsQB', function () {
            if (ExportDataArray.length != 0) {
                AccountExports.SyncAccountExports();
            }
            else {
                swal.fire({
                    title: 'Oops!',
                    text: 'Please Select a Case To Export!',
                    icon: 'warning', timer: 2000,
                    showConfirmButton: false
                });
            }
        });

        // Expand Case's Activity Table
        $(document).on('click', '.expandCaseClass', function () {
            var tr = $(this).parents('tr');//.css('background-color', 'orange');

            var row = $('#tblAccountExports').DataTable().row(tr);

            if (row.child.isShown()) {
                // The row is already open - close it
                row.child.hide();

                tr.removeClass('shown');

                $(this).children('i').removeClass('zmdi-caret-up-circle').addClass('zmdi-caret-down-circle');
                $(this).removeAttr('title').attr('title', 'Expand Case');

            }
            else {
                // Open the row
                row.child(AccountExports.innerHtml(AccountExports.iTableCounter)).show();
                tr.addClass('shown');
                if ($(this).children('i').hasClass('zmdi-caret-down-circle')) {
                    $(this).children('i').removeClass('zmdi-caret-down-circle').addClass('zmdi-caret-up-circle');
                    $(this).removeAttr('title').attr('title', 'Collapse Case');
                }

                // Get Data for Expand Table
                AccountExports.BindChildRows($(this).attr('data-CaseId'));

            }
        });

    },


    // Sync AccountExports with QuickBooks

    SyncAccountExports: function (uptoDate) {
        loader.showloader();
        var _data = "{caseDetails:" + JSON.stringify(ExportDataArray) + "}";
        ajaxrepository.callServiceWithPost('/Accounts/Export/ExportToQuickBooks', _data, AccountExports.onSyncAccountExports, AccountExports.OnError, undefined);
    },
    onSyncAccountExports: function (d, s, e) {
        if (s == "success") {
            swal.fire({
                title: "Information!",
                text: "Export has been initiated. Will Notify you once done",
                icon: "info",
                timer: 5000,
                showConfirmButton: false
            }).then((result) => {
                $("#AlreadySynced").prop("checked", true).trigger("click");
            });
            loader.hideloader();
            ExportDataCount = 0;
            ExportDataArray = [];
            $("#CheckAll").prop('checked', false);
            $(".checkBoxClass").prop('checked', false);
            checkNotification = setInterval(checkUnreadNotification, 60000);
        }
    },

    // Get and Set Case's Activity Table

    BindChildRows: function (CaseId) {
        loader.showloader();
        if (CaseId != '') {
            var _data = new Array();
            _data.push({ 'name': 'caseId', 'value': CaseId });
            ajaxrepository.callService('/Accounts/Export/GetChargesReportByCaseId', _data, AccountExports.OnBindChildRowsSuccess, this.OnError, undefined);
        }
    },
    OnBindChildRowsSuccess: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "-1") {
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else {
                _allRecord = new Array();
                $.each(d, function (index, record) {
                    _allRecord.push([
                        record.CaseChargesId,
                        record.Activity,
                        record.VenderCharge

                    ]);
                });
                $('#tblAccountExports' + AccountExports.iTableCounter).dataTable({
                    "data": _allRecord,
                    "aaSorting": [],
                    "stateSave": false,
                    "columnDefs": [{
                        "bSortable": false,
                        "mRender": function (data, type, full) {
                            return full[2] + " $";
                        },

                        "aTargets": [2]
                    },

                    ],
                    autoWidth: false,

                    deferRender: true,
                    info: false,
                    lengthChange: false,
                    ordering: false,
                    paging: false,
                    scrollX: false,
                    //  scrollY: true,
                    searching: false
                });
                //;
                nowBindedRowsCount = 0;
                $.each($('div#tblAccountExports' + AccountExports.iTableCounter + '_wrapper').find('tr'), function (i, j) {
                    nowBindedRowsCount = nowBindedRowsCount + 1;

                });
                console.log(nowBindedRowsCount);
                if (nowBindedRowsCount > 0) {
                    $('div#tblAccountExports' + AccountExports.iTableCounter + '_wrapper').find('tr:last').css('background-color', 'lightyellow');
                }
                $('div#tblAccountExports' + AccountExports.iTableCounter + '_wrapper').find('tr:last').css('background-color', 'lightyellow');

                AccountExports.iTableCounter = AccountExports.iTableCounter + 1;


            }

        }
    },


    // Token to Access QuickBooks

    GetAccessToken: function () {
        loader.showloader();
        notify.showProcessing();
        var _data = new Array();
        ajaxrepository.callService('/Accounts/Account/GetAccessToken', _data, AccountExports.onGetAccessToken, AccountExports.onErrror, undefined);
    },
    onGetAccessToken: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "-1") {
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                AccountExports.OnError()
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else {
                if (d != "") {
                    $("#btnSyncAccountExportsQB").show();
                    $("#btnConnectAccountExportsQB").hide();
                    $("#btnDeleteQBPayments").show();
                    $(".SyncControls").addClass("visible").removeClass("invisible");
                }
                else {
                    $("#btnConnecttoQuikbooks").show();
                    $("#btnSyncAccountExportsQB").hide();
                    $(".SyncControls").addClass("invisible").removeClass("visible");
                }
            }
        }
    },

    // Last Sync Date

    GetLastSyncDate: function () {
        loader.showloader();
        notify.showProcessing();
        var _data = new Array();
        ajaxrepository.callService('/Accounts/Export/GetLastSyncDate', _data, AccountExports.onGetLastSyncDate, AccountExports.onErrror, undefined);
    },
    onGetLastSyncDate: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "-1") {
            } else if (d == "-2") {
            }
            else if (d == "-5") {

                $("#LogoutModal").modal("show");
            }
            else {

                $("#lastsyncExports").html(d);
            }
        }
    },

    GetAccountExports: function (offset) {
        ExportDataCount = 0;
        ExportDataArray = [];
        ExportDataTempArray = [];
        AccountExports.recordCount = 0;

        if (ExportDataArray.length == 0) {
            $("#CheckAll").prop('checked', false);
        }
        AccountExports.oTable = $('#tblAccountExports').DataTable().destroy();

        AccountExports.oTable = $('#tblAccountExports').DataTable({
            "processing": true, // for show progress bar
            "serverSide": true, // for process server side
            responsive: true,
            language: {
                "emptyTable": "No data found"
            },
            dom: 'Bfrtip',
            "ajax": {
                url: '/Accounts/Export/GetSearchBasedExportData',
                type: 'POST',
                dataType: 'json',
                data: {
                    "IMEReportFiltter":
                        function () {
                            var search = {};
                            search.SyncedData = checkExportsSync;
                            if (AccountExports.isValidDate($("#exportFromDate").val())) {
                                search.From = new Date($("#exportFromDate").val());
                                search.FromIsValid = true;
                            } else {
                                search.FromIsValid = false;
                            }
                            if (AccountExports.isValidDate($("#exportToDate").val())) {
                                search.To = new Date($("#exportToDate").val());
                                search.ToIsValid = true;
                            } else {
                                search.ToIsValid = false;
                            }
                            search.Offset = offset;

                            var temp = new Array();
                            temp.push({ 'From': search.From, 'FromIsValid': search.FromIsValid, 'To': search.To, 'ToIsValid': search.ToIsValid, 'Offset': search.Offset, 'SyncedData': search.SyncedData, 'Offset': search.Offset });
                            console.table(temp);
                            return JSON.stringify(temp);
                        }
                },
            },
            columns: [
                { "data": "ClaimantName", "name": "ClaimantName", "autoWidth": true },
                { "data": "dateofexam", "name": "dateofexam", "autoWidth": true },
                { "data": "Doctor", "name": "Doctor", "autoWidth": true },
                { "data": "DoctorLocation", "name": "DoctorLocation", "autoWidth": true },
                { "data": "CaseChargeCount", "name": "CaseChargeCount", "autoWidth": true },
                { "data": "caseid", "name": "caseid", "autoWidth": true, "visible": true },

            ],
            buttons: [
                {
                    text: '<span><i class="fa fa-plus" aria-hidden="true"></i>PDF</span>',
                    className: 'float-right btn exportBT mr-1',

                    action: function (e, dt, node, config) {
                        AccountExports.ExportPdf();
                    }
                },
                {
                    text: '<span><i class="fa fa-plus" aria-hidden="true"></i>Excel</span>',
                    className: 'float-right btn exportBT mr-1',

                    action: function (e, dt, node, config) {

                        AccountExports.ExportExcel();
                    }
                },
            ],
            "columnDefs": [
                {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (checkExportsSync == true) {
                            if (full.ClaimantName == '') {
                                return '';
                            } else {
                                return "</a> <a href='javascript: void (0); ' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' title='Expand Case Activity Charges' ><i class='zmdi zmdi-caret-down-circle'></i></a>";
                            }
                        }
                        else {
                            if (full.ClaimantName == '') {
                                return '';
                            }
                            else {
                                if (full.ChartPrepProcessStatus == 1) {
                                    return "<a href='javascript: void (0); ' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' title='Expand Case Activity Charges' ><i class='zmdi zmdi-caret-down-circle'></i></a>  <i class='zmdi zmdi-spinner'></i>";
                                }
                                else {

                                    var tempid = (full.caseid).toString();
                                    var index = ExportDataArray.findIndex(x => x.caseIds == tempid);

                                    if (index > -1) {
                                        //return '<input type="checkbox" value="' + full.caseid + '" checked> <text>[MC]' + full.CaseNumber + '</text>';
                                        return "<a href='javascript: void (0); ' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' title='Expand Case Activity Charges' ><i class='zmdi zmdi-caret-down-circle'></i></a>  <input type='checkbox' class='checkBoxClass' value='" + full.caseid + "' data-CaseNumber = '" + full.ClaimantName + "' data-DoctorId = '" + full.caseid + "' data-CaseCharges = '" + full.CaseChargeCount + "' data-DoctorLocationId = '" + full.DoctorLocationID + "' checked>";

                                    }
                                    else {

                                        return "<a href='javascript: void (0); ' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' title='Expand Case Activity Charges' ><i class='zmdi zmdi-caret-down-circle'></i></a>  <input type='checkbox' class='checkBoxClass' value='" + full.caseid + "' data-CaseNumber = '" + full.ClaimantName + "' data-DoctorId = '" + full.caseid + "' data-CaseCharges = '" + full.CaseChargeCount + "' data-DoctorLocationId = '" + full.DoctorLocationID + "'>";
                                        //return '<input type="checkbox" value="' + full.caseid + '" > <text>[MC]' + full.CaseNumber + '</text>';
                                    }

                                    // return "<a href='javascript: void (0); ' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' title='Expand Case Activity Charges' ><i class='zmdi zmdi-caret-down-circle'></i></a>  <input type='checkbox' class='checkBoxClass' value='" + full.caseid + "' data-CaseNumber = '" + full.ClaimantName + "' data-DoctorId = '" + full.caseid + "' data-CaseCharges = '" + full.CaseChargeCount + "' data-DoctorLocationId = '" + full.DoctorLocationID + "'>";

                                }
                            }
                        }
                    },
                    "aTargets": [5]
                },
                {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (full.ClaimantName == '') {
                            return 'Total: $' + full.CaseChargeCount;
                        }
                        else {
                            return "$ " + full.CaseChargeCount;
                        }
                    },
                    "aTargets": [4]

                },

                { className: 'text-center', targets: [5] }
            ],
            "autoWidth": false,
            complete: function (data) {
                if (data.responseJSON.data != 0)
                    AccountExports.chargesCount = data.responseJSON.data[0].CaseChargeCount;
                else
                    AccountExports.chargesCount = 0;
            },
            footer: true,
            "footerCallback": function (row, data, start, end, display) {

                if (checkExportsSync == true) {
                    $('#tblAccountExports tfoot').show();
                    var api = this.api(), data;
                    var Charges = 0;
                    if (data.length > 0) {

                        Charges = data[0].TotalCharges;
                    }

                    $(api.column(4).footer()).html(
                        'Total:  $' + Charges
                    );
                }
                else {
                    $('#tblAccountExports tfoot').hide();
                }
            }
        });

        AccountExports.oTable.on('search.dt', function () {
            //number of filtered rows
            console.log(AccountExports.oTable.rows({ filter: 'applied' }).nodes().length);
            //filtered rows data as arrays
            console.log("Filtered Records");
            console.log(AccountExports.oTable.rows({ filter: 'applied' }).data());
        });

        // Check/Uncheck based on Main CheckBox
        $('#tblAccountExports').on('draw.dt', function () {

            if ($("#CheckAll").is(":checked") == true) {

                $(".checkBoxClass").prop('checked', true);

                $('#tblAccountExports tbody tr input[type=checkbox]').each(function (i) {

                    var index = ExportDataArray.findIndex(x => x.caseIds === this.defaultValue);

                    if (index == -1) {

                        $(this).prop("checked", false);

                    }

                });
            }
        });

        var index;
        var indexTemp;
        $(document).on('change', '#tblAccountExports tbody tr input[type=checkbox]', function (e) {
            if (e.target.checked) {

                index = ExportDataArray.findIndex(x => x.caseIds === this.defaultValue);

                if (index > -1) {
                    ExportDataArray.splice(index, 1);
                }
                ExportDataArray.push({ 'caseIds': this.defaultValue });

            }
            else {

                index = ExportDataArray.findIndex(x => x.caseIds === this.defaultValue);

                if (index > -1) {

                    ExportDataArray.splice(index, 1);
                }
            }

        });

    },

    //  Export Report Pdf
    ExportPdf: function () {
        loader.showloader();
        var search = {};
        search.SyncedData = checkExportsSync;
        if (AccountExports.isValidDate($("#exportFromDate").val())) {
            search.From = new Date($("#exportFromDate").val());
            search.FromIsValid = true;
        } else {
            search.FromIsValid = false;
        }
        if (AccountExports.isValidDate($("#exportToDate").val())) {
            search.To = new Date($("#exportToDate").val());
            search.ToIsValid = true;
        } else {
            search.ToIsValid = false;
        }
        if (OFFSET == 1) {
            search.Offset = 3;
        } else {
            search.Offset = 4;
        }

        search.SyncedData = checkExportsSync;
        search.DateOffset = new Date().getTimezoneOffset();
        var data = "{IMEReportFiltter:" + JSON.stringify(search) + "}";
        ajaxrepository.callServiceWithPost('/Accounts/Export/ExportReportTo', data, AccountExports.OnExportReportSuccess, AccountExports.OnError, undefined);
    },
    OnExportReportSuccess: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "-1") {
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else {
                if (d.filename != "") {
                    var win = window.open("/Case/Report/DownloadReportFile?FileName=" + d.filename + "&Option=" + 2, "_self");
                }
            }
        }
    },
    
    CheckAll: function () {
        loader.showloader();
        var search = {};

        var Customsearch = $("#tblAccountExports_filter").find("input").val();

        search.SyncedData = checkExportsSync;
        if (AccountExports.isValidDate($("#exportFromDate").val())) {
            search.From = new Date($("#exportFromDate").val());
            search.FromIsValid = true;
        } else {
            search.FromIsValid = false;
        }
        if (AccountExports.isValidDate($("#exportToDate").val())) {
            search.To = new Date($("#exportToDate").val());
            search.ToIsValid = true;
        } else {
            search.ToIsValid = false;
        }

        search.Offset = OFFSET;

        search.Search = Customsearch;


        var data = "{IMEReportFiltter:" + JSON.stringify(search) + "}";
        ajaxrepository.callServiceWithPost('/Accounts/Export/GetCaseChargesReportToBeSynced', data, AccountExports.OnCheckAllSuccess, AccountExports.OnError, undefined);
    },
    OnCheckAllSuccess: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {

            var check = $("#CheckAll").prop('checked');
            //var check = true;

            if (check == true) {

                var table = $('#tblAccountExports').DataTable();

                $(".checkBoxClass").prop('checked', check);

                var data = table
                    .rows({ filter: 'applied' })
                    .data();

                for (let index = 0; index < d.length; index++) {
                    var t = data[index];
                    var t1 = d[index].caseid;

                    if (d[index].ChartPrepProcessStatus != 1)
                        ExportDataArray.push({ 'caseIds': d[index].caseid.toString() });
                }
            }
            else {
                var data = $(".checkBoxClass").prop('checked', $(this).prop('checked'));
                ExportDataArray = [];
                ExportDataTempArray = [];
                $(".checkBoxClass").prop('checked', false);
            }
        }
    },

    // Connect To QuickBooks
    ConnectToQuickBook: function () {
        loader.showloader();
        window.location.href = "/QuickBooks/Home/Index";
    },

    isValidDate: function (s) {
        var bits = s.split('/');
        var d = new Date(bits[0] + '/' + bits[1] + '/' + bits[2]);
        return !!(d && (d.getMonth() + 1) == bits[0] && d.getDate() == Number(bits[1]));
    },

    fetchdata: function () {

        var _data = "{caseDetails:" + JSON.stringify(ExportDataArray) + "}";
        ajaxrepository.callServiceWithPost('/Accounts/Export/CheckStatusOfSync', _data, AccountExports.onfetchdata, AccountExports.OnError, undefined);

    },
    onfetchdata: function (d, s, e) {

        if (s == "success") {
            if (d == true) {
                swal.fire({
                    title: "Done!",
                    text: "Synced Record(s) of Account payable(s) with QuickBooks !",
                    icon: "success",
                    timer: 4000,
                    showConfirmButton: false
                });
            }
        }
    },
    // On Error

    OnDelay: function (d, s, e) {
        //setInterval(AccountExports.fetchdata, 180000);
        swal.fire({
            title: 'Success!',
            text: 'Process will take around 15 to minutes. please be patient and Wait.',
            icon: 'success',
            timer: 2000,
            showConfirmButton: false
        });
    },
    OnError: function (d, s, e) {
        $("#AlreadySynced").prop("checked", true).trigger("click");
        ExportDataCount = 0;
        ExportDataArray = [];
        ExportDataTempArray = [];
        loader.hideloader();
        swal.fire({
            title: 'Error!',
            text: 'Something went wrong.',
            icon: 'error',
            timer: 2000,
            showConfirmButton: false
        });
    },
}
$(document).ready(function () {

    // Initialize AccountExports Section

    AccountExports.init();

    // Get All Case's Vendors Charges


    $("#AlreadySynced").click(function () {
        ExportDataCount = 0;
        ExportDataArray = [];
        ExportDataTempArray = [];
        $("#CheckAll").hide();
        checkExportsSync = true;
        $("#exportFromDate").val("");
        $("#exportToDate").val("");
        OFFSET = 1;
        AccountExports.GetAccountExports(OFFSET);

    })

    $("#AlreadySynced").prop("checked", true).trigger("click");

    // Get All Case's Vendors Charges To Be Synced

    $("#ToBeSynced").click(function () {
        $("#CheckAll").show();
        checkExportsSync = false;
        $("#exportFromDate").val("");
        $("#exportToDate").val("");
        OFFSET = 2;
        AccountExports.GetAccountExports(OFFSET);
    });


    // Check All CheckBoxes

    $("#CheckAll").click(function () {

        AccountExports.CheckAll();

    });

});