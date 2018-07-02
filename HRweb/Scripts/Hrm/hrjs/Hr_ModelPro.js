
var jsObject = {};

var yearJson = []; //年度数据json
var dataJson = []; //表格数据json

jsObject.Setted = {

    nowYear: '',     //当前年度
    nowMonth: '',    //当前月度
    ids: '',
    colJson: [],

    init: function () {

        //初始化
        jsObject.Setted.onLoad();

        //日期更改
        $('#yearly').change(function () {

            //jsObject.Setted.InitTable();
        });

        $('#monthly1').change(function () {

            var month1 = parseInt($('#monthly1').val());
            var month2 = parseInt($('#monthly2').val());

            if (month1 > month2) {
                $('#result').html("起始月不能大于结束月..");
                $("#myModal").modal("show");
                return;
            }
            else {
                jsObject.Setted.GetColumns();
            }
        });

        $('#monthly2').change(function () {

            var month1 = parseInt($('#monthly1').val());
            var month2 = parseInt($('#monthly2').val());

            if (month1 > month2) {
                $('#result').html("起始月不能大于结束月..");
                $("#myModal").modal("show");
                return;
            }
            else {
                jsObject.Setted.GetColumns();
            }
        });

        //导出数据
        $('#btnChart').click(function () {

            $('#result').html("未更新..");
            $("#myModal").modal("show");
        });

        //导出数据
        $('#btnExp').click(function () {

            $('#result').html("未更新..");
            $("#myModal").modal("show");
            //jsObject.Setted.Export();
        });

    },

    //页面初始化
    onLoad: function () {
        
        //日期选择 初始值
        var myDate = new Date();
        var myYear = myDate.getFullYear();      //当前年度
        var myMonth = myDate.getMonth() + 1;    //当前月度

        var myMonthStr = myMonth > 9 ? myMonth : "0" + myMonth;

        //初始查询条件
        jsObject.Setted.nowYear = myYear;
        jsObject.Setted.nowMonth = myMonth;
        jsObject.Setted.SetYearly();


    },

    GetColumns: function () {

        $.ajax({
            url: "../Hrm/GetColData",
            type: 'get',
            dataType: "json",
            async: false,
            data: { startMonth: $('#monthly1').val(), endMonth: $('#monthly2').val() },
            success: function (data) {

                jsObject.Setted.colJson = data;

                jsObject.Setted.InitTable();
            }
        });
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
            //height: 500,                        //行高，如果没有设置height属性，表格自动根据记录条数觉得表格高度
            uniqueId: "no",                     //每一行的唯一标识，一般为主键列
            showToggle: true,                    //是否显示详细视图和列表视图的切换按钮
            cardView: false,                    //是否显示详细视图
            detailView: true,                   //是否显示父子表

            //表格列定义
            columns: jsObject.Setted.colJson,

            //注册加载子表的事件。注意下这里的三个参数！
            //onExpandRow: function (index, row, $detail) {
                
            //    jsObject.Setted.InitSubTable(index, row, $detail);

            //},

            url: '../Hrm/GetHr_ModelPro',        //后台数据url

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
                params.yearly = $('#yearly').val();
                params.startMonth = $('#monthly1').val();
                params.endMonth = $('#monthly2').val();

                return params;
            }
        });
    },

    //导出数据
    Export: function () {

        if (dataJson.length > 0) {

            
        }
        else {
            $('#result').html("当前无数据导出！");
            $("#myModal").modal("show");
        }
    },

    SetYearly: function () {

        $("#yearly").empty();

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
                            $("#yearly").append(options);
                        }
                        else {
                            var options = "<option value='" + yearJson[i].value + "'>" + yearJson[i].text + "</option>"
                            $("#yearly").append(options);
                        }
                    }

                    jsObject.Setted.SetMonthly1();
                }

            }
        });
    },

    SetMonthly1: function () {

        $("#monthly1").empty();

        for (var i = 1; i < 13; i++) {

            var options = "<option value='" + i + "'>" + i + "</option>"
            $("#monthly1").append(options);
        }

        jsObject.Setted.SetMonthly2();
    },

    SetMonthly2: function () {

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

        jsObject.Setted.GetColumns();
    },

    InitSubTable: function (index, row, $detail) {
        
        var parentid = row.no;
        var cur_table = $detail.html('<table></table>').find('table');
        $(cur_table).bootstrapTable({
            url: '../Hrm/GetHr_ModelPro',
            method: 'get',
            queryParams: { strParentID: parentid },
            ajaxOptions: { strParentID: parentid },
            clickToSelect: true,
            detailView: true, //父子表
            uniqueId: "MENU_ID",
            pageSize: 10,
            pageList: [10, 25],
            columns: [{
                field: 'name',
                title: '菜单URL'
            }, {
                field: 'present',
                title: '父级菜单'
            }, {
                field: 'goal',
                title: '菜单级别'
            }, ]
        });
    }


};
