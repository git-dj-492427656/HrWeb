
var jsObject = {};

var yearJson = []; //年度数据json
var dataJson = []; //表格数据json

jsObject.Setted = {

    nowYear: '',     //当前年度
    nowMonth: '',    //当前月度
    ids: '',

    init: function () {

        //初始化
        jsObject.Setted.onLoad();

        //日期更改
        $('.form_datetime').change(function () {
            jsObject.Setted.InitTable();
        });


        //生成
        $('#btnAdd').click(function () {

            jsObject.Setted.SetYearly();

            $('.msgDiv').empty();
            $('#mainData').hide();
            $('#addData').show();
        });

        //确定生成
        $('#btnSave').click(function () {

            $('.msgDiv').html('正在同步..');

            $("#btnSave").attr('disabled', true);
            $("#yearly2").attr('disabled', true);
            $("#monthly2").attr('disabled', true);

            jsObject.Setted.SynData();
        });

        $('#btnBack').click(function () {
            jsObject.Setted.Back();
        });


        //导出数据
        $('#btnExp').click(function () {

            jsObject.Setted.Export();
        });


        //批量删除
        $('#btnDel').click(function () {

            jsObject.Setted.DelList();
        });

        $('#btnOk').click(function () {

            jsObject.Setted.DelDone();
        });

    },

    //页面初始化
    onLoad: function () {

        $('#addData').hide();   //数据同步
        $('#mainData').show();  //主信息

        $(".form_datetime").datetimepicker({
            format: "yyyy-mm",      //选择后文本显示格式
            autoclose: true,
            todayBtn: true,
            todayHighlight: true,
            showMeridian: true,
            pickerPosition: "bottom-left",
            language: 'zh-CN',      //中文，需要引用zh-CN.js包
            startView: 3,          //起始选择范围：0为时间，1为日，2为月，3为年
            maxViewMode: 3,        //最大选择范围
            minView: 3            //最小选择范围
        });

        //日期选择 初始值
        var myDate = new Date();
        var myYear = myDate.getFullYear();      //当前年度
        var myMonth = myDate.getMonth() + 1;    //当前月度

        var myMonthStr = myMonth > 9 ? myMonth : "0" + myMonth;
        var endData = new Date(myYear, myMonth + 1, 1);  //后2个月，前面取当前月已+1，此处+1即可

        var endYear = endData.getFullYear();
        var endMonth = endData.getMonth() + 1;
        var endMonthStr = endMonth > 9 ? endMonth : "0" + endMonth;

        $('#startDate').val(myYear + "-" + myMonthStr);
        $('#endDate').val(endYear + "-" + endMonthStr);


        jsObject.Setted.nowYear = myYear;
        jsObject.Setted.nowMonth = myMonth;

        jsObject.Setted.InitTable(); //初始表格数据

    },

    //表格数据
    InitTable: function () {

        $('#table1').bootstrapTable('destroy');  // 销毁表格数据

        $('#table1').bootstrapTable({

            method: 'get',                      //请求方式（*）
            toolbar: '#toolbar',                //工具按钮用哪个容器
            striped: true,                      //是否显示行间隔色
            cache: false,                       //是否使用缓存，默认为true，所以一般情况下需要设置一下这个属性（*）
            pagination: false,                   //是否显示分页（*）
            sortable: true,                     //是否启用排序
            sortOrder: "asc",                   //排序方式
            //sidePagination: "server",           //分页方式：client客户端分页，server服务端分页（*）
            //pageNumber: 1,                       //初始化加载第一页，默认第一页
            //pageSize: 10,                       //每页的记录行数（*）
            //pageList: [10, 25, 50, 100],        //可供选择的每页的行数（*）
            search: true,                       //是否显示表格搜索，此搜索是客户端搜索，不会进服务端，所以，个人感觉意义不大
            contentType: "application/x-www-form-urlencoded",
            strictSearch: true,
            showColumns: false,                  //是否显示所有的列
            showRefresh: true,                  //是否显示刷新按钮
            minimumCountColumns: 2,             //最少允许的列数
            clickToSelect: true,                //是否启用点击选中行
            height: 500,                        //行高，如果没有设置height属性，表格自动根据记录条数觉得表格高度
            uniqueId: "id",                     //每一行的唯一标识，一般为主键列
            showToggle: true,                    //是否显示详细视图和列表视图的切换按钮
            cardView: false,                    //是否显示详细视图
            detailView: false,                   //是否显示父子表

            //表格列定义
            columns: [
                {
                    field: 'state',
                    align: 'center',
                    valign: 'text-bottom;',
                    checkbox: true
                },
                {
                    field: 'id',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    title: 'Id',
                    visible: false
                },
                {
                    field: 'comName',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '公司名称'
                },
                {
                    field: 'yearly',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '年 度'
                },
                {
                    field: 'monthly',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '月 度'
                },
                {
                    field: 'cxNum',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '产 线 数'
                },
                {
                    field: 'htAmount',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '合同额（万元）'
                },
                {
                    field: 'ysAmount',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '营收（万元）'
                },
                {
                    field: 'lrAmount',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '利润（万元）'
                },
                {
                    field: 'yield',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '产量（万元）'
                },
                {
                    field: 'yieEffic',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '产效（立方/人/8H）'
                },
                {
                    field: 'gjEffic',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '构件产效（立方/人/8H）'
                },
                {
                    field: 'proTeams',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '项目组数'
                },
                {
                    field: 'workDays',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: '工作天数（天）'
                }
            ],

            url: '../Hrm/GetHr_Midfact',        //后台数据url

            onLoadSuccess: function (data)         //加载数据成功事件处理
            {
                dataJson = data;
            },
            onLoadError: function () {           //加载数据失败事件处理
                alert("数据加载失败！");
            },

            //传入到后台参数
            queryParams: function (params) {

                //特别说明，返回的参数的值为空，则当前参数不会发送到服务器端
                params.startDate = $('#startDate').val();
                params.endDate = $('#endDate').val();

                return params;
            }
        });
    },


    // 批量删除
    DelList: function () {

        var rows = $("#table1").bootstrapTable('getSelections');
        if (rows.length == 0) {

            $('#result').html('请先选择要删除的记录！');
            $("#myModal").modal("show");

            return;
        }
        else {
            var ids = '';
            for (var i = 0; i < rows.length; i++) {
                ids += rows[i]['id'] + ",";
            }

            ids = ids.substring(0, ids.length - 1);
            jsObject.Setted.ids = ids;

            $('#result1').html('确定要删除吗？此操作无法撤销！');
            $("#myModal1").modal("show");

        }
    },

    //确定删除
    DelDone: function () {

        $("#myModal1").modal("toggle");

        $.ajax({
            type: 'POST',
            url: '../Hrm/DelHr_Midfact',
            data: { ids: jsObject.Setted.ids },
            success: function (data) {

                if (data == "删除成功！") {
                    jsObject.Setted.InitTable();
                }

                $('#result').html(data);
                $("#myModal").modal("show");
            }
        });
    },


    //导出数据
    Export: function () {

        if (dataJson.length > 0) {

            var _startDate = $('#startDate').val();
            var _endDate = $('#endDate').val();

            window.open("../Hrm/ExportHr_Midfact?startDate=" + _startDate + "&endDate=" + _endDate);
        }
        else {
            $('#result').html("当前无数据导出！");
            $("#myModal").modal("show");
        }
    },

    //数据同步
    SetYearly: function () {

        $("#yearly2").empty();
        $("#monthly2").empty();

        $.ajax({
            type: 'POST',
            url: '../Hrm/GetYearly',
            data: {},
            success: function (data) {
                if (data != "") {

                    yearJson = eval("(" + data + ")");

                    for (var i = 0; i < yearJson.length; i++) {
                        if (jsObject.Setted.nowYear.toString() == yearJson[i].value) {
                            var options = "<option value='" + yearJson[i].value + "' selected='selected'>" + yearJson[i].text + "</option>"
                            $("#yearly2").append(options);
                        }
                        else {
                            var options = "<option value='" + yearJson[i].value + "'>" + yearJson[i].text + "</option>"
                            $("#yearly2").append(options);
                        }
                    }
                }
                jsObject.Setted.SetMonthly();
            }
        });
    },

    SetMonthly: function () {

        $("#monthly2").empty();

        for (var i = 1; i < 13; i++) {

            if (jsObject.Setted.nowMonth == i) {
                var options = "<option value='" + i + "' selected='selected'>" + i + "</option>"
                $("#monthly2").append(options);
            }
            else {
                var options = "<option value='" + i + "'>" + i + "</option>"
                $("#monthly2").append(options);
            }
        }
    },

    SynData: function () {

        $.ajax({
            type: 'POST',
            url: '../Hrm/SynHr_Midfact',
            data: { yearly: $('#yearly2').val(), monthly: $('#monthly2').val() },
            success: function (data) {

                $('.msgDiv').html(data);

                $("#btnSave").attr('disabled', false);
                $("#yearly2").attr('disabled', false);
                $("#monthly2").attr('disabled', false);
            }
        });
    },

    //返回
    Back: function () {

        $('#addData').hide();
        $('#mainData').show();  //主信息

        jsObject.Setted.InitTable();
    }

};
