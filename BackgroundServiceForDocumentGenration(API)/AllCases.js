var oTable;
var CaseIdShow = new Array();
var CaseIdNoShow = new Array();
var CaseIdArray = new Array();
var IsSearch = false;
var AllCases = {
    iTableCounter: 1,
    table: null,
    $controlToExpand: null,
    requestForExpand: '',
    alreadyPresentSubCaseRows: 0,
    MultipleSearchClear: function () {
        $(".MultipleSearch").val("");
        //AllCases.BindReport();
        oTable.draw();
        $(".MultipleSearch").trigger('change');
    },

    docSchId: null,
    ParentDocSchId: null,
    docId: null,
    isChild: null,
    IsCancelledAppointmentCase: null,
    hasAtleastOneDefaultAppt: null,
    dataToSaveSchedule: null,
    DoctorSpecialtyId: null,
    AppDateTime: null,
    CaseStatus: null,
    RescheduleLater: null,
    CaseDocumentId: null,
    caseId: null,
    tempStatus: "",
    Show_NoShow: false,
    offset: "",
    innerHtml: function (table_id) {
        return '  <div class="table-responsive"><table id="tblcaseReport' + table_id + '" class="table table-bordered table-striped table-hover js-basic-example dataTable">\
                                <thead>\
                                    <tr>\
                                        <th>Case Id</th>\
                                            <th>\Case Number</th >\
                                            <th>\Name</th>\
                                            <th>\Case Type</th>\
                                            <th>\Service</th>\
                                            <th>\Status</th>\
                                            <th>\Date of Injury</th>\
                                            <th>\Action</th>\
                                    </tr>\
                                </thead>\
                                <tbody></tbody>\
                            </table></div>';
    },
    init: function () {


    },

    SaveSchedule: function () {

        if ($('#ScheduleLater').prop("checked") == true) {
            var AddSchedule = {};
            AddSchedule.CaseId = AllCases.CaseID;
            AddSchedule.DoctorScheduleId = AllCases.docSchId /*$("#hdnDoctorScheduleId").val();*/;
            AddSchedule.CaseStatus = true;
            AddSchedule.RescheduleLater = $('#ScheduleLater').prop("checked");
            $("#AddScheduleModalAllCase").modal('hide');
            var _data = new Array();
            var data = new Array();
            _data = "{AddSheduleModal:" + JSON.stringify(AddSchedule) + "}";
            //data.push({ 'name': 'doctorScheduleId', 'value': parseInt(docSchId) }, { 'name': 'appointmentDateTime', 'value': appointmentDateTime }, { 'name': 'comment', 'value': result.value }, { 'name': 'StatusToComplete', 'value': true });
            data.push({ 'name': 'AddSheduleModal', 'value': JSON.stringify(AddSchedule) }, { 'name': 'CaseStatusChange', 'value': true }, { 'name': 'caseId', 'value': AddSchedule.CaseId });

            AllCases.dataToSaveSchedule = _data;
            loader.showloader();
            ajaxrepository.callServiceWithPost("/Case/Doctor/AddShedule", AllCases.dataToSaveSchedule, AllCases.onSaveSchedule, AllCases.OnError, undefined);

        }
        else {
            //loader.showloader();
            var AddSchedule = {};
            AddSchedule.DoctorId = AllCases.docId;
            AddSchedule.DoctorLocationId = $("#ddlLocation").val();
            AddSchedule.CaseId = AllCases.CaseID;
            AddSchedule.description = $("#txtSchedulingNotes").val();
            AddSchedule.description2 = $("#txtSchedulingNotes2").val();
            AddSchedule.DoctorSpecialtyId = AllCases.DoctorSpecialtyId;
            AddSchedule.DoctorScheduleId = AllCases.docSchId /*$("#hdnDoctorScheduleId").val();*/;
            AddSchedule.NoShow = $('#noShow').prop("checked");

            AddSchedule.IsDefault = $('#IsDefault').prop("checked");
            AddSchedule.RescheduledDate = new Date($("#txtApptDate2").val());
            AddSchedule.ParentDoctorScheduleId = AllCases.ParentDocSchId;
            var date1_1 = $("#txtApptDate").val();
            var date1_2 = $("#txtApptTime").val();
            if ($('#ScheduleLater').prop("checked") == false) {
                AddSchedule.StartTime = dateFormat(new Date(date1_1 + ' ' + date1_2), "isoDateTime", false);
            }

            AddSchedule.SStartDate = $("#txtApptDate").val();
            AddSchedule.SStartTime = $("#txtApptTime").val();
            AddSchedule.startTimeString = $("#txtApptTime").val();
            AddSchedule.ReschedulingReason = $("#txtReschedulingReason").val();
            AddSchedule.CaseStatus = true;
            AddSchedule.IsReminder = $("#IsReminder").prop("checked");
            AddSchedule.ReminderDate = $("#txtReminder").val();

            $("#AddScheduleModalAllCase").modal('hide');
            AddSchedule.RescheduleLater = $('#ScheduleLater').prop("checked");

            var _data = new Array();
            var data = new Array();
            _data = "{AddSheduleModal:" + JSON.stringify(AddSchedule) + "}";
            //data.push({ 'name': 'doctorScheduleId', 'value': parseInt(docSchId) }, { 'name': 'appointmentDateTime', 'value': appointmentDateTime }, { 'name': 'comment', 'value': result.value }, { 'name': 'StatusToComplete', 'value': true });
            data.push({ 'name': 'AddSheduleModal', 'value': JSON.stringify(AddSchedule) }, { 'name': 'CaseStatusChange', 'value': true }, { 'name': 'caseId', 'value': AddSchedule.CaseId });

            AllCases.dataToSaveSchedule = _data;
            loader.showloader();
            ajaxrepository.callServiceWithPost('/Case/Doctor/GetCheckIfAlreadyScheduled', _data, AllCases.OnGetCheckIfAlreadyScheduled, AllCases.OnError, undefined);

        }
    },
    OnGetCheckIfAlreadyScheduled: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "0") {
                loader.showloader();
                ajaxrepository.callServiceWithPost("/Case/Doctor/AddShedule", AllCases.dataToSaveSchedule, AllCases.onSaveSchedule, AllCases.OnError, undefined);
            } else if (d == "-2") {
                AllCases.OnError();
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else if (d == "1") {
                //Schedule.ifAlreadyScheduled = true;


                swal.fire({
                    title: 'Oops..!',
                    text: 'Please reschedule as this Doctor is already having appointment for this time slot.',
                    //text: 'Please reschedule as this Doctor is already having appointment for this time slot and location combination.',
                    icon: 'warning',
                    showConfirmButton: true
                });
                //AllCases.BindReport();
                oTable.draw();
            }
        }
    },
    onSaveSchedule: function (d, s, e) {
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

                AllCases.OnError();
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else {
                swal.fire({
                    title: "Done!",
                    text: "Appointment Saved Successfully!",
                    icon: "success",
                    timer: 2000,
                    showConfirmButton: false

                }, function () {
                    AllCases.ResetForm();
                    //AllCases.BindReport();
                    oTable.draw();
                });

                if (AllCases.hasParentAppointment == true) {
                    AddClaimant.GetCaseDetailById($(".hdncaseid").val());
                }
                //AllCases.BindReport();
                oTable.draw();
            }
        }
    },
    GetAllDoctor: function (SpecialityID) {
        notify.showProcessing();
        var caseSpecialityId = '';

        try {
            caseSpecialityId = SpecialityID;
        } catch (e) {
            caseSpecialityId = $('#ddlSpecialty').val();
        }
        if (caseSpecialityId == "" || caseSpecialityId == null) {
            caseSpecialityId = 0;
        }
        if (caseSpecialityId != 0 && caseSpecialityId > 0) {
            var data = new Array();
            data.push({ 'name': 'caseSpecialityId', 'value': caseSpecialityId });
            ajaxrepository.callService('/Case/Doctor/GetAllDoctor', data, AllCases.onAllDoctorSuccess, AllCases.OnError, undefined);
        }
        else {
            $("#ddlDoctor option").remove();
        }
    },
    onAllDoctorSuccess: function (d, s, e) {
        //loader.hideloader();
        var ddDoctors = "<option value=''>-Select Doctor-</option>";
        $.each(d, function (index, Doctor) {
            ddDoctors += "<option  value='" + Doctor.DoctorId + "'>" + Doctor.DoctorId + " | " + (Doctor.lastname == null ? ' ' : Doctor.lastname) + " " + Doctor.firstname + /*" , " + Doctor.credentials + */"</option>";
        });
        $("#ddlDoctor,#ddlchangesdoctor").html(ddDoctors).val("");
        $.notifyClose();
    },
    GetAllDoctorLocationByDoctorId: function (DoctorId) {
        //loader.showloader();
        notify.showProcessing();
        var data = new Array();
        data.push({ 'name': 'DoctorId', 'value': DoctorId });
        ajaxrepository.callService('/Case/Doctor/GetDoctorLocationByLocationId', data, AllCases.onDoctorLocationSuccess, AllCases.OnError, undefined);
    },/*GetAllDoctorLocationByDoctorId*/
    onDoctorLocationSuccess: function (d, s, e) {
        //loader.hideloader();

        var ddlocation = "<option value=''>-Select Location-</option>";
        $.each(d, function (index, Location) {
            ddlocation += "<option  value='" + Location.DoctorLocationId + "'>" + Location.location + "</option>";
        });
        $("#ddlLocation").html(ddlocation).val("");
        $.notifyClose();
    },
    UpdateCaseStatusInBackground: function () {
        //ajaxrepository.callService('/Case/Case/UpdateCaseStatusInBackground', '', AllCases.onUpdateCaseStatusInBackground, AllCases.OnError, undefined);
        var offset = new Date().getTimezoneOffset();
        var data = new Array();
        data.push({ 'name': 'offset', 'value': offset });
        ajaxrepository.callService('/Case/Case/UpdateCaseStatusInBackground', data, AllCases.onUpdateCaseStatusInBackground, AllCases.OnError, undefined);
    },
    onUpdateCaseStatusInBackground: function (d, s, e) {
        //swal.fire({
        //    title: 'Done',
        //    text: 'good.',
        //    icon: 'success',
        //    timer: 1500,
        //    showConfirmButton: false
        //});
    },


    BindReport: function () {

        var UserRole = $("#UserRole").val();
        oTable = $('#tblcaseReport').DataTable({
            dom: 'Bfrtip',
            retrieve: true,
            lengthMenu: [
                [10, 25, 50, 100],
                ['10 rows', '25 rows', '50 rows', '100 rows']
            ],
            buttons: [
                //'copy', 'csv', 'excel', 'pdf', 'print'
                'pageLength'
            ],
            "language": {
                "emptyTable": "No data found",
                "search": "",
                "processing": '<i class="fa fa-spinner fa-pulse fa-3x fa-fw"></i>'
            },
            "processing": true,
            "serverSide": true, // for process server side
            "order": [[0, "desc"]],
            "ajax": {
                "url": "/Case/Case/GetCaseReport",
                "type": "POST",
                "data": {
                    "AllCaseReportFilter":
                        function () {
                            var filter = {};
                            debugger;
                            if (AllCases.tempStatus != "") {

                                filter.CaseStatusId = AllCases.tempStatus;
                            }
                            else {
                                filter.CaseStatusId = $("#ddlStatus").val() || '';
                            }

                            filter.DOB = $("#txtDOB").val();
                            //filter.Name = $("#ddlName").val() || 0;
                            filter.Name = $("#ddlName :selected").text() || '';
                            filter.ClaimNumber = $("#ddlClaimNumber :selected").text() || '';
                            filter.CaseNumber = $("#ddlCaseNumber :selected").text() || '';
                            filter.DoctorId = $("#ddlDoctorM :selected").text() || '';
                            filter.User = $("#ddlUser").val() || 0;
                            filter.Offset = AllCases.offset;
                            filter.Show_NoShow = AllCases.Show_NoShow;

                            var temp = new Array();
                            temp.push({ 'CaseStatusId': filter.CaseStatusId, 'DOB': filter.DOB, 'Name': filter.Name, 'ClaimNumber': filter.ClaimNumber, 'CaseNumber': filter.CaseNumber, 'DoctorId': filter.DoctorId, 'User': filter.User, 'Offset': filter.Offset, 'Show_NoShow': filter.Show_NoShow });

                            return JSON.stringify(temp);
                        }
                },
                "datatype": "json"
            },
            "responsive": true,
            "columns": [
                { "data": "CaseNumber", "name": "CaseNumber", "autoWidth": true },
                { "data": "Claim", "name": "Claim", "autoWidth": true },
                { "data": "NAME", "name": "NAME", "autoWidth": true },
                { "data": "DOB", "name": "DOB", "autoWidth": true },
                { "data": "CaseType", "name": "CaseType", "autoWidth": true },
                { "data": "Service", "name": "Service", "autoWidth": true },
                { "data": "CaseStatus", "name": "CaseStatus", "autoWidth": true },
                { "data": "dateofloss", "name": "dateofloss", "autoWidth": true },
                { "data": "dateofexam", "name": "dateofexam", "autoWidth": true },
                { "data": "Doctor", "name": "Doctor", "autoWidth": true },
                { "data": "DoctorLocation", "name": "DoctorLocation", "autoWidth": true },
                { "data": "Speciality", "name": "Speciality", "autoWidth": true, "visible": false },
                { "data": "Company", "name": "Company", "autoWidth": true, "visible": false },
                { "data": "caseid", "name": "caseid", "autoWidth": true, "visible": true },
                { "data": "ParentCaseId", "name": "ParentCaseId", "autoWidth": true, "visible": true },
                { "data": "SubCaseCount", "name": "SubCaseCount", "autoWidth": true, "visible": false },
                { "data": "Contact", "name": "Contact", "autoWidth": true, "visible": false },
                { "data": "Policy", "name": "Policy", "autoWidth": true, "visible": false },
                { "data": "ClientContactName", "name": "ClientContactName", "autoWidth": true, "visible": false },
                { "data": "ClientContactTel", "name": "ClientContactTel", "autoWidth": true, "visible": false },
                { "data": "Plaintiff", "name": "Plaintiff", "autoWidth": true, "visible": false },
                { "data": "Contact", "name": "Contact", "autoWidth": true, "visible": false },
                { "data": "CaseNotes", "name": "CaseNotes", "autoWidth": true, "visible": false },
                { "data": "DoctorScheduleId", "name": "DoctorScheduledId", "autoWidth": true, "visible": false },
                { "data": "IsDefaultApp", "name": "IsDefaultApp", "autoWidth": true, "visible": false },
                { "data": "ParentDoctorScheduleId", "name": "ParentDoctorScheduleId", "autoWidth": true, "visible": false },
                { "data": "SpecialityID", "name": "SpecialityID", "autoWidth": true, "visible": false },
                { "data": "RescheduleLater", "name": "ResheduleLater", "autoWidth": true, "visible": false },
                { "data": "MedicalRecordCount", "name": "MedicalRecordCount", "autoWidth": true, "visible": false },
                { "data": "WithoutMedicalRecord", "name": "WithoutMedicalRecord", "autoWidth": true, "visible": false },
                { "data": "ChartPrep", "name": "ChartPrep", "autoWidth": true, "visible": false },

            ],
            "columnDefs": [
                { type: 'date', 'targets': [3] },
                { type: 'date', 'targets': [7] },
                { type: 'date', 'targets': [8] },
                {
                    "bSortable": false,
                    "mRender": function (data, type, full) {
                        if (UserRole == "AppointmentAdmin") {
                            if (full.CaseStatus == "Complete")
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a>";
                            else
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a> <a href='javascript:void(0);' id='addSubCase' class='addSubCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.ParentCaseId == '0' ? 'inline' : 'none') + "' title='Add Sub Case' ><i class='zmdi zmdi-file-plus'></i></a> <a href='javascript:void(0);' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.SubCaseCount == '0' ? 'none' : 'inline') + "' title='Expand Case' ><i class='zmdi zmdi-caret-down-circle'></i></a> <a title='Cancel'  href='javascript:void(0)' class='btnCancelAppointmentAllCase' data-DocSchId=" + full.DoctorScheduleId + " data-ParentDocSchId=" + full.ParentDoctorScheduleId + " data-isChild=" + (full.Doctor != null ? 1 : 0) + " data-appointmentDateTime='" + full.dateofexam + "' data-IsDefault='" + full.IsDefaultApp + "' ><i class='zmdi zmdi-calendar-close'></i></a> ";

                        }
                        else if (UserRole == "CaseManager") {
                            if (full.CaseStatus == "Complete")
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a>";
                            else
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a> <a href='javascript:void(0);' id='addSubCase' class='addSubCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.ParentCaseId == '0' ? 'inline' : 'none') + "' title='Add Sub Case' ><i class='zmdi zmdi-file-plus'></i></a> <a href='javascript:void(0);' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.SubCaseCount == '0' ? 'none' : 'inline') + "' title='Expand Case' ><i class='zmdi zmdi-caret-down-circle'></i></a>  ";

                        }
                        else if (UserRole == "Doctor") {
                            return "";
                        }
                        else {
                            if (full.CaseStatus == "Complete")
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a> <a href='javascript:void(0);' id='addSubCase' class='addSubCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.ParentCaseId == '0' ? 'inline' : 'none') + "' title='Add Sub Case' ><i class='zmdi zmdi-file-plus'></i></a> <a href='javascript:void(0);' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.SubCaseCount == '0' ? 'none' : 'inline') + "' title='Expand Case' ><i class='zmdi zmdi-caret-down-circle'></i></a> <a title='Preview Document'  href='javascript:void(0)' class='btPreviewDocument' data-CaseID='" + full.caseid + "'><i class='zmdi zmdi-eye'></i></a>";
                            else
                                return "<input type='hidden' value='" + full.caseid + "' class='hdncaseid'/><a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank'><i class='zmdi zmdi-comment-edit'></i></a> <a href='javascript:void(0);' id='addSubCase' class='addSubCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.ParentCaseId == '0' ? 'inline' : 'none') + "' title='Add Sub Case' ><i class='zmdi zmdi-file-plus'></i></a> <a href='javascript:void(0);' id='expandCase' class='expandCaseClass' data-CaseId='" + full.caseid + "' style='display :" + (full.SubCaseCount == '0' ? 'none' : 'inline') + "' title='Expand Case' ><i class='zmdi zmdi-caret-down-circle'></i></a> <a title='Cancel'  href='javascript:void(0)' class='btnCancelAppointmentAllCase' data-DocSchId=" + full.DoctorScheduleId + " data-ParentDocSchId=" + full.ParentDoctorScheduleId + " data-isChild=" + (full.Doctor != null ? 1 : 0) + " data-appointmentDateTime='" + full.dateofexam + "' data-IsDefault='" + full.IsDefaultApp + "' ><i class='zmdi zmdi-calendar-close'></i></a> <a title='Reschedule'  href='javascript:void(0)' class='btAddAppointment' data-SpecialityID='" + full.SpecialityID + "' data-CaseID='" + full.caseid + "' data-AppDateTime='" + full.dateofexam + "' data-CaseStatus='" + full.CaseStatus + "' data-RescheduleLater='" + full.RescheduleLater + "'><i class='zmdi zmdi-calendar'></i></a> <a title='Medical Record'  href='javascript:void(0)' class='btAddMedicalRecord' data-CaseID='" + full.caseid + "'><i class='zmdi zmdi-hospital'></i></a> <a title='Preview Document'  href='javascript:void(0)' class='btPreviewDocument' data-CaseID='" + full.caseid + "'><i class='zmdi zmdi-eye'></i></a> <a title='Skip Chart Prep'  href='javascript:void(0)' class='btSkip' data-CaseID='" + full.caseid + "'><i class='zmdi zmdi-skip-next'></i></a>";

                        }
                    },

                    "aTargets": [13]
                },
                {
                    "mRender": function (data, type, full) {
                        if (UserRole == "AppointmentAdmin") {
                            return full.CaseNumber;
                        }
                        else if (UserRole == "Doctor") {
                            return full.CaseNumber;
                        }
                        else {

                            //if (full.ParentCaseId == 0) {
                            //    return '<input type="checkbox" value="' + full.caseid + '" > <text>[MC]' + full.CaseNumber + '</text>';
                            //}
                            //else {
                            //    return '<input type="checkbox" value="' + full.caseid + '" > <text>[SC]' + full.CaseNumber + '</text>';
                            //}

                            var tempid = (full.caseid).toString();
                            var index = CaseIdArray.indexOf(tempid);
                            if (full.ParentCaseId == 0) {
                                if (index > -1) {
                                    return '<input type="checkbox" value="' + full.caseid + '" checked> <text>[MC]' + full.CaseNumber + '</text>';
                                }
                                else {
                                    return '<input type="checkbox" value="' + full.caseid + '" > <text>[MC]' + full.CaseNumber + '</text>';
                                }
                            }
                            else {
                                if (index > -1) {
                                    return '<input type="checkbox" value="' + full.caseid + '" checked> <text>[SC]' + full.CaseNumber + '</text>';
                                }
                                else {
                                    return '<input type="checkbox" value="' + full.caseid + '" > <text>[SC]' + full.CaseNumber + '</text>';
                                }
                            }

                        }
                    },

                    "aTargets": [0]
                },
                {
                    "bSortable": false,
                    "width": "5%",
                    "mRender": function (data, type, full) {
                        if ((full.dateofexam == "" || full.dateofexam == null) || (full.CaseStatus != "Awaiting Appointment" && full.CaseStatus != "Awaiting Scheduling")) {
                            return '';
                        }
                        else {
                            return '<div style="text-align: center;display: ' + (AllCases.CheckIfAppointmentDateGreaterToNow(full.dateofexam) == true ? "none" : "in-line") + '"><input title="Show" class="show" id = "Show_NoShow' + full.caseid + '" name = "Show_NoShow' + full.caseid + '" type = "radio" style = "transform:scale(1.2);" data-CaseId="' + full.caseid + '"/>  &nbsp;  <input title="No-Show" class="noshow" id="Show_NoShow' + full.caseid + '" name="Show_NoShow' + full.caseid + '" type="radio" value="NoShow" class="NoShow" style = "transform:scale(1.2);" data-CaseId="' + full.caseid + '"/>  &nbsp;  <input title="Not Applicable" class="notapplicable" id="Show_NoShow' + full.caseid + '" name="Show_NoShow' + full.caseid + '" type="radio" value="NotApplicable" style = "transform:scale(1.2);" data-CaseId="' + full.caseid + '"/></div>';
                        }
                    },

                    "aTargets": [14]
                },
                {
                    "aTargets": [10], "visible": true
                },
                {
                    "aTargets": [11], "visible": false
                },
                {
                    "aTargets": [12], "visible": false
                },
                {

                    "mRender": function (data, type, full) {

                        if (full.CaseStatus == "Complete") {
                            return "<a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank' class='statuscolor cancelledstatus'>" + full.NAME + "</a>";

                        }
                        else if ((full.MedicalRecordCount > 0 || full.WithoutMedicalRecord == true) && full.ChartPrep == true) {
                            return "<a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank' class='statuscolor greenstatus'>" + full.NAME + "</a>";

                        }
                        else if (full.MedicalRecordCount > 0 || full.WithoutMedicalRecord == true) {
                            return "<a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank' class='statuscolor bluestatus'>" + full.NAME + "</a>";

                        }
                        else if (full.ChartPrep == true) {
                            return "<a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank' class='statuscolor magentastatus'>" + full.NAME + "</a>";

                        }
                        else {
                            return "<a  href='/Case/Case/AddCase?ID=" + full.caseid + "&CaseNumber=" + full.CaseNumber + "' title='Edit Case' target='_blank' class='statuscolor'>" + full.NAME + "</a>";

                        }


                    },

                    "aTargets": [2]
                }

            ],

            autoWidth: false


        });

        $(".MultipleSearch").trigger('change');

        $(document).on('click', '#btGenerateDoc', function () {
            Swal.fire({
                title: 'You are sure you want to generate document ?',
                text: "Are You Sure?",
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#3085d6',
                cancelButtonColor: '#d33',
                confirmButtonText: 'Yes',
                cancelButtonText: 'No',

            }).then((result) => {
                if (result.value) {

                    if (CaseIdArray.length != 0)
                        var data = {
                            BatchCaseId: CaseIdArray,
                            offset: 2,
                            DateOffset: new Date().getTimezoneOffset()
                        };
                    var _data = new Array();
                    _data = "{AddDocumentModals:" + JSON.stringify(data) + "}";
                    loader.showloader();
                    ajaxrepository.callServiceWithPostAndCache("/Case/Docments/CallChartPrepWebApi", _data, AllCases.onApiHitSuccess, AllCases.OnError, undefined);

                }
            });
        });

        var index;
        $(document).on('change', '#tblcaseReport tbody tr input[type=checkbox]', function (e) {

            if (e.target.checked) {

                $("#btGenerateDoc").show();

                index = CaseIdArray.indexOf(this.defaultValue);

                if (index > -1) {
                    CaseIdArray.splice(index, 1);

                }

                CaseIdArray.push($(this).val());
            }
            else {

                index = CaseIdArray.indexOf(this.defaultValue);

                if (index > -1) {

                    CaseIdArray.splice(index, 1);

                }

            }
            if (CaseIdArray.length == 0) {

                $("#btGenerateDoc").hide();

            }

            if (CaseIdArray.length >= 1000) {

                $('#tblcaseReport tbody tr input[type=checkbox]').not($("#tblcaseReport tbody tr input:checked")).each(function (i) {

                    $(this).attr("disabled", true);

                });
            }
            else {

                $('#tblcaseReport tbody tr input[type=checkbox]').each(function (i) {

                    $(this).attr("disabled", false);

                });
            }

        });

        $('#tblcaseReport').on('draw.dt', function () {

            if (CaseIdArray.length >= 1000) {

                $('#tblcaseReport tbody tr input[type=checkbox]').not($("#tblcaseReport tbody tr input:checked")).each(function (i) {

                    $(this).attr("disabled", true);

                });
            }
            else {

                $('#tblcaseReport tbody tr input[type=checkbox]').each(function (i) {

                    $(this).attr("disabled", false);

                });
            }
        });

        $('#tblcaseReport tbody').on('change', 'tr input[type=radio]', function () {
            if (this.checked) {
                if (this.title == "No-Show") {
                    var index = CaseIdShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdShow.splice(index, 1);
                    }
                    index = CaseIdNoShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdNoShow.splice(index, 1);
                    }
                    CaseIdNoShow.push(parseInt(this.dataset.caseid));
                }
                else if (this.title == "Show") {
                    var index = CaseIdNoShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdNoShow.splice(index, 1);
                    }
                    index = CaseIdShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdShow.splice(index, 1);
                    }
                    CaseIdShow.push(parseInt(this.dataset.caseid));
                }
                else {
                    var index = CaseIdNoShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdNoShow.splice(index, 1);
                    }
                    index = CaseIdShow.indexOf(parseInt(this.dataset.caseid));
                    if (index > -1) {
                        CaseIdShow.splice(index, 1);
                    }
                }
            }
            return false;
        });


        //$('#tblcaseReport').on('draw.dt', function () {

        var table = $('#tblcaseReport').DataTable();
        if (UserRole == "AppointmentAdmin") {
            table.columns([0, 1, 2, 5, 6, 8, 9, 10, 11]).visible(true);
            table.columns([3, 4, 7, 12]).visible(false);
        }
        else if (UserRole == "Doctor") {
            table.columns([0, 1, 2, 5, 6, 8, 9, 10, 11]).visible(true);
            table.columns([3, 4, 7, 12, 14]).visible(false);
        }
        else {
            if ($('#ddlStatus :selected').text() == "Awaiting Scheduling" || $('#ddlStatus :selected').text() == "Chart preparation") {
                table.columns([3, 5, 6, 7]).visible(false);
                table.columns([10, 11, 12]).visible(true);
            }
            else {
                table.columns([3, 11, 12]).visible(false);
                table.columns([5, 6, 7, 10]).visible(true);
            }
        }
        var check = $('#tblcaseReport thead tr:eq(1) th');
        //var checkRow = $('#tblcaseReport thead tr').length;
        if (check.length == 0) {

            $('#tblcaseReport thead tr').clone(true).appendTo('#tblcaseReport thead');
            $('#tblcaseReport thead tr:eq(1) th').each(function (i) {

                if ($(this).text() != "Action" && $(this).text() != "Show/No-Show") {
                    $(this).html('<input type="text" placeholder="Search" class="form-control MultipleSearch"  style= "margin-left:-3%;width:80%"/>');

                    //For remove sorting from datacolumn search
                    $(this).removeClass("sorting sorting_asc sorting_desc");

                    //For multiple search(each column search)
                    $('input', this).on('keyup change', function () {

                        oTable.columns(i).search($(this).val().trim());

                        oTable.draw();//Hit server

                    });

                    $('input', this).on('click', function (e) {
                        e.stopPropagation();
                    });
                }
                else if ($(this).text() == "Action") {
                    //$(this).html('<a  href="javascript:void(0);" onclick="AllCases.MultipleSearchClear()" title="Clear Search"><i class="zmdi zmdi-close-circle zmdi-hc-2x pl-3"></i></a>');
                    $(this).html('<a type="button" class="btn btn-round btn-primary btn-sm" href="javascript:void(0);" onclick="AllCases.MultipleSearchClear()" title="Clear Search">Clear</i></a>');
                }
                else {
                    $(this).html('<div style="text-align:center;"><button type="button" id="btShow_NoShow" style="margin-top:28px;" title="Show/NoShow" class="btn btn-primary btn-sm btn-round">UpdateStatus</button></div>');
                }
            });
        }
        else {

            var temp = 0;

            $('#tblcaseReport thead tr').each(function (index) {
                if (temp != 0)
                    $(this).remove();
                temp++;
            });

            var checkRowss = $('#tblcaseReport thead tr').length;
            $('#tblcaseReport thead tr').clone(true).appendTo('#tblcaseReport thead');
            $('#tblcaseReport thead tr:eq(1) th').each(function (i) {

                if ($(this).text() != "Action" && $(this).text() != "Show/No-Show") {
                    $(this).html('<input type="text" placeholder="Search" class="form-control MultipleSearch"  style= "margin-left:-3%;width:80%"/>');

                    //For remove sorting from datacolumn search
                    $(this).removeClass("sorting sorting_asc sorting_desc");

                    //For multiple search(each column search)
                    $('input', this).on('keyup change', function () {

                        oTable.columns(i).search($(this).val().trim());

                        oTable.draw();//Hit server

                    });

                    $('input', this).on('click', function (e) {
                        e.stopPropagation();
                    });
                }
                else if ($(this).text() == "Action") {
                    //$(this).html('<a  href="javascript:void(0);" onclick="AllCases.MultipleSearchClear()" title="Clear Search"><i class="zmdi zmdi-close-circle zmdi-hc-2x pl-3"></i></a>');
                    $(this).html('<a type="button" class="btn btn-round btn-primary btn-sm" href="javascript:void(0);" onclick="AllCases.MultipleSearchClear()" title="Clear Search">Clear</i></a>');
                }
                else {
                    $(this).html('<div style="text-align:center;"><button type="button" id="btShow_NoShow" style="margin-top:28px;" title="Show/NoShow" class="btn btn-primary btn-sm btn-round">UpdateStatus</button></div>');
                }
            });
        }
    },
    onApiHitSuccess: function (d, s, e) {
        CaseIdArray = [];
        $("#btGenerateDoc").hide();
        //AllCases.BindReport();
        oTable.draw();
        loader.hideloader();
        if (s == 'success') {
            Swal.fire({
                title: 'Please Be Patient',
                text: "Will notify you once Done.",
                icon: 'info',
                timer: 2000,
                showConfirmButton: false
            });

            checkNotificationChart = setInterval(checkUnreadNotification, 60000);

        }
        else {
            swal.fire({
                title: 'Error!',
                text: 'Something went wrong.',
                icon: 'error',
                timer: 2000,
                showConfirmButton: false
            });
        }
    },
    onAddDocumentSuccess: function (d, s, e) {
        if (s == "success") {
            if (d == "-1") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });

            }
            else if (d == "-5") {
                loader.hideloader();
                $("#LogoutModal").modal("show");
            }
            else {
                var data = {
                    offset: 0,
                    CaseIdShow: CaseIdNoShow
                };
                var p = "{statusDetails:" + JSON.stringify(data) + "}"
                ajaxrepository.callServiceWithPost('/Case/Docments/GetSpecificDocument', p, AllCases.onDownloadSuccess, AllCases.OnError, undefined);
            }
        }
    },
    downloadAll: function (urls) {
        var link = document.createElement('a');

        link.setAttribute('download', null);
        link.style.display = 'none';

        document.body.appendChild(link);

        for (var i = 0; i < urls.length; i++) {
            link.setAttribute('href', urls[i]);
            link.click();
        }

        document.body.removeChild(link);
    },
    onDownloadSuccess: function (d, s, e) {
        if (s == 'success') {
            if (d != '0') {
                Swal.fire({
                    title: 'Thanks for your patience so far...',
                    text: "Download will complete in some moments",
                    icon: 'info',
                    timer: 2000,
                    showConfirmButton: false
                });
                var links = [];
                var dlength = d.length;
                for (var i = 0; i < dlength; i++) {
                    links.push("/Case/Docments/DownloadAzureLink?FilePath=" + d[i].pathname + "&FileName=" + d[i].DocmentOrgName + "&CaseId=" + d[i].CaseId);
                }
                AllCases.downloadAll(links);
                CaseIdNoShow = [];
                CaseIdShow = [];
                oTable.draw();
                loader.hideloader();

            }
            else {
                CaseIdNoShow = [];
                CaseIdShow = [];
                oTable.draw();
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            }
        }
        else {
            CaseIdNoShow = [];
            CaseIdShow = [];
            oTable.draw();
            loader.hideloader();
            swal.fire({
                title: 'Error!',
                text: 'Something went wrong.',
                icon: 'error',
                timer: 2000,
                showConfirmButton: false
            });
        }
    },

    AddSubCase: function (caseId) {
        var _data = new Array();

        _data.push({ 'name': 'caseId', 'value': caseId });
        ajaxrepository.callService('/Case/Case/AddSubCase/', _data, AllCases.onAddSubCaseSuccess, AllCases.OnError, undefined);
    },

    onAddSubCaseSuccess: function (d, s, e) {
        if (s == 'success') {
            if (d != '0') {

                $('#addSubCaseModal').modal('hide');

                swal.fire({
                    title: 'Done!',
                    text: 'Sub case added successfully!!',
                    icon: 'success',
                    timer: 2000,
                    showConfirmButton: false
                });
                //-Initiate Grid for sub case grid-
                AllCases.InitiateSubCaseGridAfterAddSubCase(AllCases.$controlToExpand);
                //-Initiate Grid for sub case grid-

                //AllCases.BindReport();
            }
            else {
                $('#addSubCaseModal').modal('hide');
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrongSub Case not added.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });

            }
        }
        else {
            $('#addSubCaseModal').modal('hide');
            swal.fire({
                title: 'Error!',
                text: 'Something went wrong.',
                icon: 'error',
                timer: 2000,
                showConfirmButton: false
            });
        }
    },
    //BindChildRows: function (CaseId, control) {
    BindChildRows: function (CaseId) {
        //;
        loader.showloader();
        //console.log(control);
        if (CaseId != '') {
            var _data = new Array();
            _data.push({ 'name': 'caseId', 'value': CaseId }, { 'name': 'casestatusid', 'value': $("#ddlStatus").val() });
            ajaxrepository.callService('/Case/Case/GetSubCasesReportByCaseId', _data, AllCases.OnBindChildRowsSuccess, this.OnError, undefined);
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
                        record.caseid,
                        record.CaseNumber,
                        record.NAME,
                        record.CaseType,
                        record.Service,
                        record.CaseStatus,
                        record.dateofloss,
                        record.ParentCaseId

                    ]);
                });
                $('#tblcaseReport' + AllCases.iTableCounter).dataTable({
                    "data": _allRecord,
                    "aaSorting": [],
                    "stateSave": false,
                    "columnDefs": [{
                        "bSortable": false,
                        "mRender": function (data, type, full) {
                            //console.log(full);
                            return "<a  href='/Case/Case/AddCase?ID=" + full[0] + "&CaseNumber=" + full[1] + "' title='Edit Case'><i class='zmdi zmdi-comment-edit'></i></a>";

                            //return '<a data-toggle="tooltip" data-placement="bottom" title="Edit" target="_blank" class="action-icons" onclick="ContractorList.editContractorClick(\'' + full[1] + '\')"   ><i class="glyphicon glyphicon-edit"></i></a> <a data-toggle="tooltip" data-placement="bottom" title="View" class="action-icons" onclick="ContractorList.GetContractorDetail(\'' + full[0] + '\')"><i class="glyphicon glyphicon-open-file"></i></a><a data-toggle="tooltip" data-placement="bottom" title="View Comments" class="action-icons" onclick="ContractorList.getcomments(' + full[1] + ')"><i class="glyphicon glyphicon-comment"></i></a>';

                        },

                        "aTargets": [7]
                    },

                    ],
                    autoWidth: false,

                    deferRender: true,
                    info: false,
                    lengthChange: false,
                    ordering: false,
                    paging: false,
                    scrollX: false,
                    scrollY: true,
                    searching: false
                });
                //;
                nowBindedRowsCount = 0;
                $.each($('div#tblcaseReport' + AllCases.iTableCounter + '_wrapper').find('tr'), function (i, j) {
                    nowBindedRowsCount = nowBindedRowsCount + 1;

                });
                console.log(AllCases.alreadyPresentSubCaseRows);
                console.log(nowBindedRowsCount);
                if (nowBindedRowsCount > 0 && AllCases.alreadyPresentSubCaseRows == 0) {
                    $('div#tblcaseReport' + AllCases.iTableCounter + '_wrapper').find('tr:last').css('background-color', 'lightyellow');
                }
                if (AllCases.alreadyPresentSubCaseRows < nowBindedRowsCount && AllCases.alreadyPresentSubCaseRows > 0) {
                    $('div#tblcaseReport' + AllCases.iTableCounter + '_wrapper').find('tr:last').css('background-color', 'lightyellow');
                }
                alreadyPresentSubCaseRows = 0;
                AllCases.iTableCounter = AllCases.iTableCounter + 1;


            }

        }
    },
    InitiateSubCaseGridAfterAddSubCase: function (c) {
        var tr = $controlToExpand;

        var row = $('#tblcaseReport').DataTable().row(tr);

        row.child(AllCases.innerHtml(AllCases.iTableCounter)).show();

        tr.addClass('shown');
        tr.find('a.expandCaseClass').css('display', 'inline');
        if ($(this).children('i').hasClass('zmdi-caret-down-circle')) {
            $(this).children('i').removeClass('zmdi-caret-down-circle').addClass('zmdi-caret-up-circle');
            $(this).removeAttr('title').attr('title', 'Collapse Case');
        }
        AllCases.BindChildRows(tr.find('a.addSubCaseClass').attr('data-CaseId'));
    },
    EditCase: function (caseId, CaseNumber) {
        loader.showloader();
        window.location = href = '/Case/Case/AddCase?ID=' + caseId + "&CaseNumber=" + CaseNumber;
    },
    ReschedulingAppointment: function (SpecialityID, RescheduleLater) {
        $(".title").html("Schedule");
        if (RescheduleLater == "true") {
            AllCases.GetAllDoctor(SpecialityID);

            $('#noShow').prop('checked', false);
            $('.noShow').hide();
            AllCases.Disablefields(false);
            $('.RescheduleDateDiv').addClass("d-none");
            AllCases.docSchId = 0;
            $("#AddScheduleModalAllCase").modal("show");
            AllCases.ResetForm();
            $("#txtApptDate").prop('disabled', true);
            $("#txtApptTime").prop('disabled', true);
            $("#txtApptDate2").prop('disabled', true);
            $("#ddlLocation").prop('disabled', true);
            $(".ddlDoctorDisable").prop('disabled', true);
            $("#DoctorSpecialtyId").prop('disabled', true);
            $("#txtSchedulingNotes").prop('disabled', true);
            $("#txtReschedulingReason").prop('disabled', true);
            $('#IsDefault').prop('checked', true);
            $('#IsDefault').prop('disable', false);
            $('#ScheduleLater').prop('checked', true);
        }
        else {
            AllCases.GetAllDoctor(SpecialityID);

            $("#AddScheduleModalAllCase").modal("show");
            $('#noShow').prop('checked', false);
            $('.noShow').hide();
            AllCases.Disablefields(false);
            $('.RescheduleDateDiv').addClass("d-none");
            AllCases.docSchId = 0;
            AllCases.ResetForm();
            AllCases.hasParentAppointment = false;
            $('#ScheduleLater').prop('checked', false);
            $(".ddlDoctorDisable").prop('disabled', false);
            $("#txtSchedulingNotes").prop('disabled', false);
            $("#txtReschedulingReason").prop('disabled', false);
            $('#IsDefault').prop('checked', true);
            $('#IsDefault').prop('disable', false);

        }
    },
    AddMedicalRecord: function () {
        Swal.fire({
            title: 'You are about to Add a Medical Document for this case?',
            text: "Do You wish to continue?",
            icon: 'warning',
            showCancelButton: true,
            showCloseButton: true,
            confirmButtonColor: '#3085d6',
            cancelButtonColor: '#d33',
            confirmButtonText: 'Yes, I do!',
            cancelButtonText: 'No, Proceed without Medical Document',
        }).then((result) => {
            if (result.value) {
                $("#AddDocumentModal").modal("show");
                $(".title").html("Add Medical Document");
                $("#dropdocfile").val(null);
                $("#ddlDocumentClass").prop("disabled", true);
                $("#ddlDocumentTypes").prop("disabled", true);
                AllCases.CaseDocumentId = 0;
                $("#DropFile").show();
            }
            else if (result.dismiss == 'cancel') {
                var offset = new Date().getTimezoneOffset();
                var data = new Array();
                data.push({ 'name': 'caseid', 'value': AllCases.caseId });
                loader.showloader();
                ajaxrepository.callService('/Case/Report/UpdateWithoutMedicalRecord', data, AllCases.onChangeStatus, AllCases.OnError, undefined);
            }
        });

    },
    StatusToAwaitingAccounting: function (CaseIdShow) {
        var data = {
            offset: 1,
            CaseIdShow: CaseIdShow
        };
        var p = "{statusDetails:" + JSON.stringify(data) + "}"
        loader.showloader();
        ajaxrepository.callServiceWithPost('/Case/Report/UpdateBatchStatus', p, AllCases.AddDocument, AllCases.OnError, undefined);
    },
    onAwaitingAccountingChange: function (d, s, e) {
        if (s == "success") {
            if (d == "-1") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });

            }
            else if (d == "-5") {
                loader.hideloader();
                $("#LogoutModal").modal("show");
            }
            else {
                var data = {
                    offset: 2,
                    CaseIdShow: CaseIdNoShow
                };
                var p = "{statusDetails:" + JSON.stringify(data) + "}"
                loader.showloader();
                ajaxrepository.callServiceWithPost('/Case/Report/UpdateBatchStatus', p, AllCases.onAddDocumentSuccess, AllCases.OnError, undefined);
            }
        }
    },
    AddDocument: function (d, s, e) {
        CaseIdShow = [];
        if (s == "success") {
            if (d == "-1") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Please contact to administrator.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });
            } else if (d == "-2") {
                loader.hideloader();
                swal.fire({
                    title: 'Error!',
                    text: 'Something went wrong.',
                    icon: 'error',
                    timer: 2000,
                    showConfirmButton: false
                });

            }
            else if (d == "-5") {
                loader.hideloader();
                $("#LogoutModal").modal("show");
            }
            else {
                if (CaseIdNoShow.length != 0) {
                    var NewDocument = {};
                    NewDocument.CaseDocumentId = 0;
                    NewDocument.TemplateTypeID = 2;
                    NewDocument.TemplateId = 21;
                    NewDocument.Reference = "NoShow Ltr - NSL";
                    NewDocument.DocumentClassId = "";
                    NewDocument.DocumentStatusID = "";
                    NewDocument.Exrefernce = "";
                    NewDocument.DocumentTypeId = 11;
                    NewDocument.DocumentName = "";
                    NewDocument.DocumentOrgName = "";
                    NewDocument.DocumentExt = ".DOCX";
                    NewDocument.ServiceId = $("#ddlDocumentService").val();
                    NewDocument.TemplateName = "NSL";
                    NewDocument.CaseIdShow = CaseIdNoShow;

                    var _data = new Array();
                    var _data = "{AddDocumentModal:" + JSON.stringify(NewDocument) + "}";
                    loader.showloader();
                    ajaxrepository.callServiceWithPost("/Case/Docments/AddTemplate", _data, AllCases.onAwaitingAccountingChange, AllCases.OnError, undefined);
                }
                else {
                    loader.hideloader();
                    swal.fire({
                        title: 'Success!',
                        text: 'Status is updated.',
                        icon: 'info',
                        timer: 2000,
                        showConfirmButton: false
                    });
                    //AllCases.BindReport();
                    oTable.draw();
                }
            }
        }
    },
    AddDirectDocument: function (CaseIdNoShow) {
        CaseIdShow = [];
        if (CaseIdNoShow.length != 0) {
            var NewDocument = {};
            NewDocument.CaseDocumentId = 0
            NewDocument.TemplateTypeID = 2;
            NewDocument.TemplateId = 21;
            NewDocument.Reference = "NoShow Ltr - NSL";
            NewDocument.DocumentClassId = "";
            NewDocument.DocumentStatusID = "";
            NewDocument.Exrefernce = "";
            NewDocument.DocumentTypeId = 11;
            NewDocument.DocumentName = "";
            NewDocument.DocumentOrgName = "";
            NewDocument.DocumentExt = ".DOCX";
            NewDocument.ServiceId = $("#ddlDocumentService").val();
            NewDocument.TemplateName = "NSL";
            NewDocument.CaseIdShow = CaseIdNoShow;

            var _data = new Array();
            var _data = "{AddDocumentModal:" + JSON.stringify(NewDocument) + "}";
            loader.showloader();
            ajaxrepository.callServiceWithPost("/Case/Docments/AddTemplate", _data, AllCases.onAwaitingAccountingChange, AllCases.OnError, undefined);
        }
        else {
            loader.hideloader();
            swal.fire({
                title: 'Success!',
                text: 'Status is updated.',
                icon: 'info',
                timer: 2000,
                showConfirmButton: false
            });
            //AllCases.BindReport();
            oTable.draw();
        }
    },
    onChangeStatus: function (d, s, e) {
        loader.hideloader();
        CaseIdShow = [];
        CaseIdNoShow = [];
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
                swal.fire({
                    title: "Success!",
                    text: "Case Status has been updated!",
                    icon: "success",
                    timer: 2000,
                    showConfirmButton: false,
                }).then((result) => {
                    // AllCases.BindReport();
                    oTable.draw();
                });
            }
        }
    },
    SaveDocument: function () {
        $("#AddDocumentModal").modal("hide");
        loader.showloader();
        var uploadedFiles = $('#dropdocfile')[0].files;
        var formData = new FormData();
        if (uploadedFiles.length > 0) {

            for (var i = 0; i < uploadedFiles.length; i++) {
                formData.append(uploadedFiles[i].name, uploadedFiles[i]);
            }

        }
        formData.append('CaseDocumentId', 0);
        formData.append('CaseId', AllCases.caseId);
        formData.append('DocumentClassId', 13);
        formData.append('DocumentTypeId', 1);
        formData.append('TemplateTypeID', 2);
        formData.append('TemplateId', "");
        formData.append('Reference', "");
        formData.append('DocumentStatusID', "");
        formData.append('Exrefernce', "");
        formData.append('DocumentName', "");
        formData.append('DocumentOrgName', "");
        formData.append('DocumentExt', "");
        formData.append('ServiceId', "");

        if ($("#ddlTemplate").val() != "") {
            formData.append('TemplateName', "");
        }
        var _data = new Array();
        var _data = "{AddDocumentModal:" + JSON.stringify(formData) + "}";
        loader.showloader();
        ajaxrepository.callServiceWithAttachment("/Case/Docments/AddDocument", formData, AllCases.onSaveDocument, AllCases.OnError, undefined);
    },
    onSaveDocument: function (d, s, e) {
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
                //AllCases.BindReport();
                oTable.draw();
                swal.fire({
                    title: "Done!",
                    text: "Medical Record has been Saved Successfully!",
                    icon: "success",
                    timer: 2000,
                    showConfirmButton: false
                });
            }
        }
    },
    ReschedulingAppointmentSave: function (SpecialityID) {
        var ScheduleLater = $('#ScheduleLater').prop("checked");
        if (ScheduleLater == true) {
            AllCases.SaveSchedule();
        }
        else {
            if ($('#frmScheduleAllCase').valid()) {
                var IsDate = Date.parse($("#txtApptDate2").val())
                var date1_1 = $("#txtApptDate").val();
                var date1_2 = $("#txtApptTime").val();
                if (IsDate) {
                    var date1 = new Date(date1_1 + date1_2);
                    var date2 = new Date($("#txtApptDate2").val());
                    var dateSum = new Date(date1_1 + date1_2);

                    if (dateSum > date1 || date1_2 == date2.getTime()) {
                        swal.fire({
                            title: 'Oops..!',
                            text: 'Rescheduled Appt. Date & Time should be greater than original Appt. Date & time.',
                            icon: 'warning',
                            timer: 4500,
                            showConfirmButton: true
                        });
                    } else {
                        AllCases.SaveSchedule();
                    }
                }
                else {
                    AllCases.SaveSchedule();
                }
            }
        }

    },
    ResetForm: function () {
        $("#txtApptDate").val("");
        $("#txtApptTime").val("");
        $("#txtApptDate2").val("");
        $("#ddlLocation").val("");
        $("#ddlDoctor").val("").trigger('change.select2').length;
        $("#DoctorSpecialtyId").val("").trigger('change.select2').length;
        $("#txtSchedulingNotes").val("");
        $('#IsDefault').prop('checked', true);
        $('#txtReschedulingReason').val("");
        $('#IsDefault').prop('disabled', true);
        $('#IsReminder').prop('checked', false);
        $("#txtReminder").prop("disabled", true);
        $("#txtReminder").val("");
        //if (AllCases.hasAtleastOneDefaultAppt) {
        //    $('#IsDefault').prop('disabled', true);
        //} else {
        //    $('#IsDefault').prop('disabled', false);
        //}

        var validator = $("#frmScheduleAllCase").validate();
        validator.resetForm();
    },
    Disablefields: function (value) {

        if (value) {
            $("#txtApptDate").prop('disabled', true);
            $("#txtApptTime").prop('disabled', true);
            $("#ddlDoctor").prop('disabled', true);
            $("#ddlLocation").prop('disabled', true);
            $("#DoctorSpecialtyId").prop('disabled', true);
        }
        else {
            $("#txtApptDate").prop('disabled', false);
            $("#txtApptTime").prop('disabled', false);
            $("#ddlDoctor").prop('disabled', false);
            $("#ddlLocation").prop('disabled', false);
            $("#DoctorSpecialtyId").prop('disabled', false);
            //$("#IsDefault").prop('disabled', false);
            //$("#spnDateOfLoss").removeAttr('readonly');

        }
        //$('#noShow').val(this.checked);
    },
    formvalidation: function () {
        $('#frmScheduleAllCase').validate({
            rules: {
                'txtApptDate': {
                    required: true
                },
                'txtApptTime': {
                    required: true
                },
                'txtApptDate2': {
                    required: true
                },
                'ddlDoctor': {
                    required: true
                },
                'ddlLocation': {
                    required: true
                },
                'DoctorSpecialtyId': {
                    required: true
                },
                'txtReschedulingReason': {
                    required: true
                },
                'txtSchedulingNotes': {
                    required: function (element) {
                        if ($("#noShow").is(':checked')) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    }
                }
            },
            messages: {
                'txtApptDate': {
                    required: "Please enter Appt Date!"
                },
                'txtApptTime': {
                    required: "Please enter Appt Time!"
                },
                'txtApptDate2': {
                    required: "Please enter Rescheduled Date & Time!"
                },
                'ddlDoctor': {
                    required: "Please select Doctor first!"
                },
                'DoctorSpecialtyId': {
                    required: "Please select Doctor Specialty!"
                },
                'ddlLocation': {
                    required: "Please select Location first!"
                },
                'txtSchedulingNotes': {
                    required: "Scheduling Notes are required!"
                },
                'txtReschedulingReason': {
                    required: "Rescheduling Reason is required!"
                }
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
    CancelAppointment: function (docSchId, ParentDocSchId, IsChild, appointmentDateTime, IsDefault) {
        //alert(IsDefault);
        loader.hideloader();
        Swal.fire({
            title: 'You are about to Cancel this Case\'s Default Appointment. It will take this case to "Complete" Status.',
            text: "Are You Sure?",
            icon: 'warning',
            input: 'text',
            inputPlaceholder: 'Reason for Cancelling',
            showCancelButton: true,
            confirmButtonColor: '#3085d6',
            cancelButtonColor: '#d33',
            confirmButtonText: 'Yes, Cancel it!',
            cancelButtonText: 'Close',
            inputValidator: (value) => {
                return !value && 'Please provide us some Reason!'
            }
        }).then((result) => {
            if (result.value) {
                if (appointmentDateTime != "null") {
                    Cancelled_Appointment = true;
                    AllCases.docSchId = docSchId;
                    AllCases.ParentDocSchId = ParentDocSchId;
                    AllCases.isChild = IsChild;
                    if (appointmentDateTime != null) {
                        AllCases.IsCancelledAppointmentCase = true;
                    }
                    else {
                        AllCases.IsCancelledAppointmentCase = false;
                    }
                    var data = new Array();
                    data.push({ 'name': 'doctorScheduleId', 'value': parseInt(docSchId) }, { 'name': 'appointmentDateTime', 'value': appointmentDateTime }, { 'name': 'comment', 'value': result.value }, { 'name': 'StatusToComplete', 'value': true });
                    loader.showloader();
                    ajaxrepository.callService('/Case/Doctor/CancelDoctorScheduled', data, AllCases.onCancelAppointmentSuccess, AllCases.OnError, undefined);

                }
                else {
                    Swal.fire('No Default Appointment is found for this Case');
                }
            } else {

            }
        });
    },
    onCancelAppointmentSuccess: function (d, s, e) {
        loader.hideloader();
        if (s == "success") {
            if (d == "0") {
                swal.fire({
                    title: 'Error!',
                    text: 'Appointment is not Cancelled',
                    icon: 'error',
                    showConfirmButton: true
                });
            } else if (d == "-2") {
                AllCases.OnError();
            }
            else if (d == "-5") {
                $("#LogoutModal").modal("show");
            }
            else if (d == "1") {

                swal.fire({
                    title: 'Done!',
                    text: 'Appointment is Cancelled successfully.',
                    icon: 'success',
                    timer: 4000,
                    showConfirmButton: false
                }).then((result) => {
                    //AllCases.BindReport();
                    oTable.draw();
                });
            }
            else if (d == "3") {
                swal.fire({
                    title: 'Oops..!',
                    text: 'To Cancel an appointment, the appointment date time should be greated than now.',
                    icon: 'warning',
                    showConfirmButton: true
                }).then((result) => {
                });
            }
        }
    },
    CheckIfAppointmentDateGreaterToNow: function (appointmentDateTime) {
        var CurrentDate = new Date();
        appointmentDateTime = new Date(appointmentDateTime);

        if (appointmentDateTime > CurrentDate) {
            return true;
        } else {
            return false;
        }
    },
    GetInformationForSideBar: function (caseid) {

        var caseId = caseid;
        var data = new Array();
        if (caseId != 0) {
            data.push({ 'name': 'caseId', 'value': caseId });
            $("#clientInformation").addClass("d-none");
            $("#loadingOnDelayClientInfo").removeClass("d-none");
            ajaxrepository.callService('/Case/Case/GetInformation', data, AllCases.onGetInformationForSideBarSuccess, AllCases.OnError, undefined);
        }
    },
    onGetInformationForSideBarSuccess: function (d, s, e) {

        var temp = '<hr><div class="informationSideBar">'
            + '<span><span style="color: rgb(0, 188, 212)"><strong><u>Claimant:- </u></strong></span><br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Name: </span>' + d[0].NAME + ' <br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Contact: </span>' + d[0].Contact + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)"><strong><u>Client:- </u></strong></span><br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Name: </span>' + d[0].ClientContactName + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Contact: </span>' + d[0].ClientContactTel + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Company: </span>' + d[0].Company + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)"><strong><u>Case And Service:- </u></strong></span><br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Case Type: </span>' + d[0].CaseType + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Speciality: </span>' + d[0].Speciality + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">WCB: </span>' + d[0].Policy + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Service: </span>' + d[0].Service + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)"><strong><u>Doctor:- </u></strong></span><br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Name: </span>' + d[0].Doctor + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Location: </span>' + d[0].DoctorLocation + '<br></span>'
            + '<span><span style="color: rgb(0, 188, 212)">Contact: </span>' + d[0].Phone + '<br></span>'
            + '<span style="color: rgb(0, 188, 212)">Notes: </span><span title="' + d[0].CaseNotes + '">' + d[0].CaseNotes + '<br></span>'
            + '</div>'

        $("#clientInformation").removeClass("d-none");
        $("#loadingOnDelayClientInfo").addClass("d-none");
        $("#clientInformation").html(temp);
    },

    OnError: function (d, s, e) {
        CaseIdArray = [];
        loader.hideloader();
        swal.fire({
            title: 'Error!',
            text: 'Something went wrong.',
            icon: 'error',
            timer: 1500,
            showConfirmButton: false
        });
    },
}
$(document).ready(function () {
    AllCases.init();

    var t1 = $('#tblcaseReport').dataTable();
    //oTable = $('#tblcaseReport').dataTable();

    t1.on('mouseover', 'tr', function () {
        $("#tblcaseReport tr").not(this)
            .removeClass("tooltip-highlight")
            .removeAttr('title')
            .removeClass("tooltip-text")
            .removeAttr('data-toggle')
            .removeAttr('data-placement');

        if ($(this).hasClass("tooltip-highlight")) {

            $(this).removeClass("tooltip-highlight");
            $(this).removeAttr('title');
            $(this).removeClass("tooltip-text");
            $(this).removeAttr('data-toggle');
            $(this).removeAttr('data-placement');

        }
        else {

            var data = $('#tblcaseReport').DataTable().rows(this).data()[0]
            if (data != undefined) {
                var title = `CaseId: ${data.caseid}\n\nClaim Number #: ${data.Claim}\n\nClaimant Name: ${data.NAME}\n\nClaimant Contact: ${data.Contact}\n\nClientContact Name: ${data.ClientContactName}\n\nClientContactTel #: ${data.ClientContactTel}\n\nWCB #: ${data.Policy}\n\nDateOfLoss: ${data.dateofloss}\n\nPlaintiff: ${data.Plaintiff}\n\nDoctor Note: ${data.CaseNotes}`;
                this.setAttribute('title', title);
                this.classList.add("tooltip-text");
                this.classList.add("tooltip-highlight");
                this.setAttribute('data-toggle', "tooltip");
                this.setAttribute('data-placement', "bottom");
            }
        }
    });

    t1.on('click', 'tr', function () {

        var data = $('#tblcaseReport').DataTable().rows(this).data()[0];
        $("#tblcaseReport tr").not(this)
            .removeClass("rowHighlight")

        if ($(this).hasClass("rowHighlight")) {

            $("#clientInformation").html("");
            $(this).removeClass("rowHighlight");

        }
        else {
            if (data != undefined) {

                AllCases.GetInformationForSideBar(data.caseid);
                this.setAttribute('class', 'rowHighlight');

            }
        }
    });

    //Custom search for datatable(Hit on click of search button)
    $('#btSearch').on('click', function () {

        AllCases.BindReport();
    });

    $("#SaveDocument").click(function () {
        var temp = $("#disp_tmp_path").val();
        if ($("#disp_tmp_path").val() == "" || $("#disp_tmp_path").val() == null) {
            swal.fire({
                title: 'Error!',
                text: 'Please Upload file first!!.',
                icon: 'warning',
                timer: 2000,
                showConfirmButton: false
            });
        }
        else {
            if ($('#AddDocument').valid()) {
                AllCases.SaveDocument();
            }
        }
    });
    $(document).on("click", "#btShow_NoShow", function () {
        console.log(CaseIdShow)
        console.log(CaseIdNoShow)
        if (CaseIdShow.length != 0) {
            AllCases.StatusToAwaitingAccounting(CaseIdShow);
        }
        else if (CaseIdNoShow.length != 0) {
            AllCases.AddDirectDocument(CaseIdNoShow);
        }
        else {
            swal.fire({
                title: 'Warning!',
                text: 'Please Select a Case First!',
                icon: 'warning',
                timer: 2000,
                showConfirmButton: false
            });
        }
    });

    $("#dropdocfile").change(function (event) {
        var filePath = event.target.files[0];
        $("#disp_tmp_path").val(filePath);
    });

    $('#txtApptDate').datetimepicker({
        language: 'en',
        format: 'm/d/Y',
        timepicker: false,
        minDate: 0,
        scrollMonth: false,
        scrollInput: false
    });

    $('#ddlDoctor').change(function () {
        if ($(this).val() != "") {
            AllCases.docId = $(this).val();
            AllCases.GetAllDoctorLocationByDoctorId($(this).val());
        } else {
            var ddlocation = "<option value=''>-Select Location-</option>";
            $("#ddlLocation").html(ddlocation).val("");
        }
    });

    $('#txtApptTime').datetimepicker({
        language: 'en',
        datepicker: false,
        formatTime: 'h:i A',
        step: 15,
        format: 'h:i A',
        minTime: '06:15',
        maxTime: '22:15',
        validateOnBlur: false
    });

    $('#txtReminder').datetimepicker({
        language: 'en',
        format: 'm/d/Y H:m A',
        //timepicker: false,
        minDate: 0,
        validateOnBlur: false
    });

});
