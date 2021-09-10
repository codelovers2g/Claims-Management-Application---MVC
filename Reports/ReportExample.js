var LstDailyStatusDataReport = [];
var DoctorDailyStatus = {
    iTableCounter: 1,
    table: null,
    otable: null,
    recordCount: 0,
    sortBy: '',
    sortOrder: '',
    search: '',
    init: function () {
        $("#ReportMenu").click();
        $('#DailyStatusAnchor').css('color', '#00bcd4');
        $('#tblIMEDataReportReport').dataTable();
        DoctorDailyStatus.formvalidation();

        $("#btnSearchDailyStatusReport").click(function () {
            var Date = $("#txtForDate").val();
            var DoctorId = $("#ddlDoctor").val();
            var LocationId = $("#ddlLocation").val();

            if (DoctorId == "") {
                $("#btnSubmitDailyStatus").hide();
            }
            else {
                $("#btnSubmitDailyStatus").show();
            }
            if (Date != "" || DoctorId != "" || LocationId != "") {
                if ($('#frmDailyStatus').valid()) {
                    if ($("#txtForDate").val().indexOf('_') > -1) {
                        alert("Invalid Date");
                    } else {
                        DoctorDailyStatus.BindReport();
                    }
                }
            }
            else {
                swal.fire("Please Select Atleast One Field!");
            }

        });
    },

    ExportWord: function () {
        loader.showloader();
        var search = {};

        if (DoctorDailyStatus.isValidDate($("#txtForDate").val())) {
            search.For = new Date($("#txtForDate").val());
            search.FromText = $("#txtForDate").val();
            search.FromIsValid = true;
        } else {
            search.FromIsValid = false;
        }

        search.SortBy = DoctorDailyStatus.sortBy;
        search.SortOrder = DoctorDailyStatus.sortOrder;
        search.Search = DoctorDailyStatus.search;
        search.ToIsValid = false;
        search.DoctorId = $("#ddlDoctor").val();
        search.LocationId = $("#ddlLocation").val() || 0;
        search.isDailyStatusReport = true;
        search.isAppointmentLogReport = false;
        search.Offset = 2;
        search.DateOffset = new Date().getTimezoneOffset();
        var data = "{IMEReportFiltter:" + JSON.stringify(search) + "}";
        ajaxrepository.callServiceWithPost('/Case/Report/ExportReportTo', data, DoctorDailyStatus.onExportReportSuccess, DoctorDailyStatus.OnError, undefined);
    },
    onExportReportSuccess: function (d, s, e) {
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
    ExportExcel: function () {
        loader.showloader();
        var search = {};

        if(DoctorDailyStatus.isValidDate($("#txtForDate").val())) {
            search.For = new Date($("#txtForDate").val());
            search.FromText = $("#txtForDate").val();
            search.FromIsValid = true;
        } else {
            search.FromIsValid = false;
        }

        search.SortBy = DoctorDailyStatus.sortBy;
        search.SortOrder = DoctorDailyStatus.sortOrder;
        search.Search = DoctorDailyStatus.search;
        search.ToIsValid = false;
        search.DoctorId = $("#ddlDoctor").val();
        search.LocationId = $("#ddlLocation").val() || 0;
        search.isDailyStatusReport = true;
        search.isAppointmentLogReport = false;
        search.Offset = 2;
        search.DateOffset = new Date().getTimezoneOffset();
        var data = "{IMEReportFiltter:" + JSON.stringify(search) + "}";
        ajaxrepository.callServiceWithPost('/Case/Report/ExportExcelReportTo', data, DoctorDailyStatus.onExportExcelReportSuccess, DoctorDailyStatus.OnError, undefined);
    },
    onExportExcelReportSuccess: function (d, s, e) {
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
                    var win = window.open("/Case/Report/DownloadReportFile?FileName=" + d.filename + "&Option=" + 3, "_self");
                }
            }
        }
    },
    isValidDate: function (s) {
        var bits = s.split('/');
        var d = new Date(bits[0] + '/' + bits[1] + '/' + bits[2]);
        return !!(d && (d.getMonth() + 1) == bits[0] && d.getDate() == Number(bits[1]));
    },
    BindReport: function () {
        
        DoctorDailyStatus.recordCount = 0;
        //DoctorDailyStatus.recordCount = d.length;

        var oTable = $('#tblDailyStatus').dataTable();
        oTable.fnClearTable();
        oTable.fnDestroy();
        oTable = $('#tblDailyStatus').dataTable({
            //dom: 'lBfrtip',
            "processing": true, // for show progress bar
            "serverSide": true, // for process server side
            lengthMenu: [
                [10, 25, 50, 100],
                ['10 rows', '25 rows', '50 rows', '100 rows']
            ],
            responsive: true,
            language: {
                "emptyTable": "No data found"
            },
            dom: 'Bfrtip',
            "ajax": {
                url: '/Case/Report/GetDoctorDailyStatusReport',
                type: 'POST',
                dataType: 'json',
                data: {
                    "IMEReportFiltter":
                        function () {
                            var search = {}

                            if (DoctorDailyStatus.isValidDate($("#txtForDate").val())) {
                                search.For = new Date($("#txtForDate").val());
                                search.FromText = $("#txtForDate").val();
                                search.FromIsValid = true;
                            }
                            else
                            {
                                search.FromIsValid = false;
                            }

                            search.DoctorId = $("#ddlDoctor").val() || 0;
                            search.LocationId = $("#ddlLocation").val() || 0;
                            search.Offset = 1;
                            search.DateOffset = new Date().getTimezoneOffset();
                            var temp = new Array();
                            temp.push({ 'For': search.For, 'FromIsValid': search.FromIsValid, 'FromText': search.FromText, 'DoctorId': search.DoctorId, 'LocationId': search.LocationId, 'Offset': search.Offset, 'DateOffset': search.DateOffset });
                            console.table(temp);
                            return JSON.stringify(temp);
                        }
                },
                complete: function (data) {
                    if (data.responseJSON.data != 0) {

                        DoctorDailyStatus.recordCount = data.responseJSON.data[0].recordsTotal;
                        DoctorDailyStatus.sortBy = data.responseJSON.data[0].SortBy;
                        DoctorDailyStatus.sortOrder = data.responseJSON.data[0].SortOrder;
                        DoctorDailyStatus.search = data.responseJSON.data[0].Search;

                    }
                    else
                        DoctorDailyStatus.recordCount = 0;

                    
                },
            },
            columns: [
                { "data": "AppointmentTimeString", "name": "AppointmentTimeString", "autoWidth": true },
                { "data": "ClaimantFullName", "name": "ClaimantFullName", "autoWidth": true },
                { "data": "casenumber", "name": "casenumber", "autoWidth": true },
                { "data": "Service", "name": "Service", "autoWidth": true },
                { "data": "ShowsUp", "name": "ShowsUp", "autoWidth": true },
                { "data": "NoShow", "name": "NoShow", "autoWidth": true },
                { "data": "UnableToExamine", "name": "UnableToExamine", "autoWidth": true, "width": "15%" },
                { "data": "Comments", "name": "Comments", "autoWidth": true, "width": "20%" },
            ],
            buttons: [

                'pageLength',

                {
                    text: '<span>PDF</span>',
                    className: 'float-right btn exportBT mr-1',

                    action: function (e, dt, node, config) {
                        if (DoctorDailyStatus.recordCount > 0) {

                            DoctorDailyStatus.ExportWord();
                        } else {
                            swal.fire({
                                title: 'Oops..!',
                                text: 'Data required for Export.',
                                icon: 'waring',
                                timer: 2000,
                                showConfirmButton: false
                            });
                        }
                    }
                },
                {
                    text: '<span>Excel</span>',
                    className: 'float-right btn exportBT mr-1',

                    action: function (e, dt, node, config) {
                        DoctorDailyStatus.ExportExcel();
                    }
                },
            ],
            "columnDefs": [
                {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (type === 'export') {

                            if (full.ShowsUp == true)
                            {
                                return "X";
                            }
                            else
                            {
                                return " ";
                            }

                        }
                        else {

                            var tempid = (full.DoctorScheduleId).toString();
                            var index = LstDailyStatusDataReport.findIndex(x => x.DoctorScheduleId == tempid);

                            if (index > -1)
                            {
                                return "<input type='checkbox' class='p-2 doctorSchedule ShowsUp'" + (LstDailyStatusDataReport[index].ShowsUp == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)'  style='width: 17px;height:17px;' id='ShowsUp" + full.DoctorScheduleId + "' /> <input type='hidden'  value= " + full.DoctorScheduleId + " data-docScheduleId='" + full.DoctorScheduleId + "' />  ";
                            }
                            else
                            {
                                return "<input type='checkbox' class='p-2 doctorSchedule ShowsUp'" + (full.ShowsUp == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)'  style='width: 17px;height:17px;' id='ShowsUp" + full.DoctorScheduleId + "' /> <input type='hidden'  value= " + full.DoctorScheduleId + " data-docScheduleId='" + full.DoctorScheduleId + "' />  ";
                            }

                        }
                    },
                    "aTargets": [4]
                }, {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (type === 'export') {
                            if (full.NoShow == true) {
                                return "X";
                            }
                            else
                            {
                                return " ";
                            }
                        }
                        else {


                            var tempid = (full.DoctorScheduleId).toString();
                            var index = LstDailyStatusDataReport.findIndex(x => x.DoctorScheduleId == tempid);

                            if (index > -1) {

                                return "<input type='checkbox' class='p-2 doctorSchedule NoShow '" + (LstDailyStatusDataReport[index].NoShow == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='width: 17px;height:17px;' id='NoShow" + full.DoctorScheduleId + "' />  ";
                            }
                            else
                            {
                                return "<input type='checkbox' class='p-2 doctorSchedule NoShow '" + (full.NoShow == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='width: 17px;height:17px;' id='NoShow" + full.DoctorScheduleId + "' />  ";
                            }

                        }
                    },
                    "aTargets": [5]
                }, {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (type === 'export') {
                            if (full.UnableToExamine == true) {
                                return "X";
                            }
                            else {
                                return " ";
                            }
                        } else {

                            var tempid = (full.DoctorScheduleId).toString();
                            var index = LstDailyStatusDataReport.findIndex(x => x.DoctorScheduleId == tempid);

                            if (index > -1) {

                                return "<input type='checkbox' class='p-2 doctorSchedule UnableToExamine'" + (LstDailyStatusDataReport[index].UnableToExamine == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='width: 17px;height:17px;' id='UnableToExamine" + full.DoctorScheduleId + "' />  ";

                            }
                            else {

                                return "<input type='checkbox' class='p-2 doctorSchedule UnableToExamine'" + (full.UnableToExamine == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='width: 17px;height:17px;' id='UnableToExamine" + full.DoctorScheduleId + "' />  ";

                            }

                        }
                    },
                    "aTargets": [6]
                },
                {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (type === 'export') {

                            return full.Comments;
                        } else {
                            var datafull;
                            if (full.Comments == null) {
                                datafull = "";
                            } else {
                                datafull = full.Comments;
                            }

                            var tempid = (full.DoctorScheduleId).toString();
                            var index = LstDailyStatusDataReport.findIndex(x => x.DoctorScheduleId == tempid);

                            if (index > -1) {

                                //return "<input type='checkbox' class='p-2 doctorSchedule ShowsUp'" + (LstDailyStatusDataReport[index].ShowsUp == true ? "checked" : "false") + " onChange='DoctorDailyStatus.onChangeCheckbox(this)'  style='width: 17px;height:17px;' id='ShowsUp" + full.DoctorScheduleId + "' /> <input type='hidden'  value= " + full.DoctorScheduleId + " data-docScheduleId='" + full.DoctorScheduleId + "' />  ";
                                return "<input type='text' data-comment='" + full.UnableToExamine + "' value='" + LstDailyStatusDataReport[index].Comments + "'   title='comment' onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='border-radius: 3px;  border: 1px solid #888; padding: 2px; width: 100%;' maxlength='200'/>";

                            }
                            else {

                                return "<input type='text' data-comment='" + full.UnableToExamine + "' value='" + datafull + "'   title='comment' onChange='DoctorDailyStatus.onChangeCheckbox(this)' style='border-radius: 3px;  border: 1px solid #888; padding: 2px; width: 100%;' maxlength='200'/>";

                            }

                        }
                    },
                    "aTargets": [7]
                },
                { className: 'text-center', targets: [4, 5, 6] },
              
            ],
            "autoWidth": true
        });

    },

    SaveDailyStatus: function () {
        DailyStatusReportModel = new Object();
        DailyStatusReportModel.DoctorId = $("#ddlDoctor").val();
        DailyStatusReportModel.DoctorLocationId = $("#ddlLocation").val();
        DailyStatusReportModel.LstDailyStatusData = LstDailyStatusDataReport;
        
        if (DailyStatusReportModel.LstDailyStatusData.length > 0) {
            loader.showloader();
            var _data = new Array();
            var _data = "{objDailyStatusReportModel:" + JSON.stringify(DailyStatusReportModel) + "}";
            ajaxrepository.callServiceWithPost('/Case/Report/SaveDailyStatusReport', _data, DoctorDailyStatus.onSaveDailyStatusSuccess, DoctorDailyStatus.onServicError, undefined);

        }
        else
        {
            swal.fire({
                title: 'Oops..!',
                text: 'Please  select atleast one checkbox!',
                icon: 'warning',
                timer: 2000,
                showConfirmButton: false
            });
        }
    },
    onSaveDailyStatusSuccess: function (d, s, e) {
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
                    timer: 1500,
                    showConfirmButton: false
                });
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else {
                swal.fire({
                    title: 'Done!',
                    text: 'Saved Successfully!',
                    icon: 'success',
                    timer: 2000,
                    showConfirmButton: false
                }).then((result) => {
                    DoctorDailyStatus.BindReport();
                    LstDailyStatusDataReport = [];
                });
            }
        }
    },
    onChangeCheckbox: function (control) {

        var ClassName = control.className;
        if (ClassName.indexOf('NoShow') != -1) {
            ClassName = "NoShow";
        }
        else if (ClassName.indexOf('ShowsUp') != -1) {
            ClassName = "ShowsUp";
        }
        else if (ClassName.indexOf('UnableToExamine') != -1) {
            ClassName = "UnableToExamine";
        }
        //control.id.replace(/[^0-9]/g, '')
        if ($(control).is(':checked')) {

            if (ClassName == "ShowsUp") {
                $(control).parents('tr').children('td:nth-child(6)').find('input[type=checkbox], select').prop('checked', false);
                $(control).parents('tr').children('td:nth-child(7)').find('input[type=checkbox], select').prop('checked', false);
            }
            else if (ClassName == "NoShow") {
                $(control).parents('tr').children('td:nth-child(5)').find('input[type=checkbox], select').prop('checked', false);
                $(control).parents('tr').children('td:nth-child(7)').find('input[type=checkbox], select').prop('checked', false);
            }
            else if (ClassName == "UnableToExamine") {
                $(control).parents('tr').children('td:nth-child(5)').find('input[type=checkbox], select').prop('checked', false);
                $(control).parents('tr').children('td:nth-child(6)').find('input[type=checkbox], select').prop('checked', false);
            }

        }
        if (!$(control).is(':checked')) {
        }

        var comment = $(control).closest('tr').children('td:nth-child(8)').find('input[type=text], select').val();

        //$(DoctorDailyStatus.otable.fnGetNodes()).each(function (index, element) {
 
        var ShowsUp = $(control).closest('tr').children('td:nth-child(5)').find('input[type=checkbox], select').prop('checked');
        var NoShow = $(control).closest('tr').children('td:nth-child(6)').find('input[type=checkbox], select').prop('checked');
        var UnableToExamine = $(control).closest('tr').children('td:nth-child(7)').find('input[type=checkbox], select').prop('checked');

        //if (ShowsUp || NoShow || UnableToExamine) {
        var docSchId = $(control).closest('tr').find('input[type=hidden], select').attr('data-docScheduleId');
        var comment = $(control).closest('tr').children('td:nth-child(8)').find('input[type=text], select').val();

        var DailyStatusReportListingModel = {};
        DailyStatusReportListingModel.ShowsUp = ShowsUp;
        DailyStatusReportListingModel.NoShow = NoShow;
        DailyStatusReportListingModel.UnableToExamine = UnableToExamine;
        DailyStatusReportListingModel.Comments = comment;
        DailyStatusReportListingModel.DoctorScheduleId = $(control).closest('tr').find('input[type=hidden], select').val();

        index = LstDailyStatusDataReport.findIndex(x => x.DoctorScheduleId === DailyStatusReportListingModel.DoctorScheduleId);

        if (index > -1) {

            LstDailyStatusDataReport.splice(index, 1);

        }
        LstDailyStatusDataReport.push(DailyStatusReportListingModel);
    },
    GetAllDoctorLocationByDoctorId: function (DoctorId) {
        notify.showProcessing();
        var data = new Array();
        data.push({ 'name': 'DoctorId', 'value': DoctorId });
        ajaxrepository.callService('/Case/Doctor/GetDoctorLocationByLocationId', data, DoctorDailyStatus.onDoctorLocationSuccess, DoctorDailyStatus.OnError, undefined);
    },
    onDoctorLocationSuccess: function (d, s, e) {
        var ddlocation = "<option value=''>-Select Location-</option>";
        $.each(d, function (index, Location) {
            ddlocation += "<option  value='" + Location.DoctorLocationId + "'>" + Location.location + "</option>";
        });
        $("#ddlLocation").html(ddlocation).val("");
        $.notifyClose();
    },
    formvalidation: function () {
        $('#frmDailyStatus').validate({
            rules: {
                'DoctorId': {
                    //required: true
                    required: false
                },
                'ddlLocation': {
                    //required: true
                    required: false
                },
                'txtForDate2': {
                    //required: true
                    required: false
                },
            },
            messages: {
                'txtForDate2': {
                    required: "Please enter Date first!"
                },
                'DoctorId': {
                    required: "Please select Doctor first!"
                },
                'ddlLocation': {
                    required: "Please select Location first!"
                },
            },
            highlight: function (input) {
                $(input).parents('.form-line').addClass('error');
            },
            unhighlight: function (input) {
                $(input).parents('.form-line').removeClass('error');
            },
            errorPlacement: function (error, element) {
                $(element).parents('.form-group,.cot-validation,.fordateerror').append(error);
            }
        });
    },
    OnError: function (d, s, e) {
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
    DoctorDailyStatus.init();
    $("#ddlDoctor").change(function () {
        if ($(this).val() != "") {
            DoctorDailyStatus.GetAllDoctorLocationByDoctorId($(this).val());
        } else {
            var ddlocation = "<option value=''>-Select Location-</option>";
            $("#ddlLocation").html(ddlocation).val("");
        }
    });

    $(document).on('change', '#ddlDoctor', function () {
        $(this).valid();
    });

    DoctorDailyStatus.otable = $('#tblDailyStatus').dataTable();
    
    $("#btnSubmitDailyStatus").click(function () {
        //if ($('#AddPayment').valid()) {
        if (DoctorDailyStatus.recordCount > 0) {
            //if ($(".doctorSchedule :checked", oTable.fnGetNodes()).length > 0) {
            DoctorDailyStatus.SaveDailyStatus();

        } else {
            swal.fire({
                title: 'Oops..!',
                text: 'Data Required',
                icon: 'warning',
                timer: 2000,
                showConfirmButton: false
            });
        }
    });
    var d = new Date();
    var month = d.getMonth() + 1;
    var day = d.getDate();
    var currentDate = (month < 10 ? '0' : '') + month + '/' + (day < 10 ? '0' : '') + day + '/' + d.getFullYear();
    $('#txtForDate').val(currentDate);
});