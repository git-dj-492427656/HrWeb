
var jsObject = {};

var dataJson; //表格数据json

jsObject.Setted = {

    init: function () {

        //初始化
        jsObject.Setted.onLoad();

        //日期更改
        $('.form_datetime').change(function () {
            jsObject.Setted.InitTable();
        });

        //生成
        $('#btnAdd').click(function () {

            $('#result').html("无数据源Bais接口！");
            $("#myModal").modal("show");
        });

        //导入数据
        $('#btnImp').click(function () {

            $('#mainData').hide();
            $('#addData').hide();
            $('#import').show();

        });

        $('#btnBack2').click(function () {
            jsObject.Setted.Back();
        });


        //导出数据
        $('#btnExp').click(function () {

            jsObject.Setted.Export();
        });


        //确定删除
        $('#btnDel').click(function () {

            jsObject.Setted.DelDone();
        });

    },

    //页面初始化
    onLoad: function () {

        $('#addData').hide();   //数据同步
        $('#import').hide();    //上传文件
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
                    field: 'baisComName',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    sortable: true,
                    title: 'Bais公司名称'
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
                    title: '产量（万立方）'
                },
                {
                    field: 'id',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '160',
                    title: '操 作',
                    events: operateEvents,
                    formatter: jsObject.Setted.operateFormatter //自定义方法，添加操作按钮
                }
            ],

            url: '../Hrm/GetHr_Bais_xmjz',        //后台数据url

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

    //表格按钮赋值
    operateFormatter: function (value, row, index) {
        return [
               '<a class="RoleOfB btn btn-danger btn-xs" href="#"><i class="fa fa-trash"></i>&nbsp;删除</a>'
        ].join('');
    },

    // 删除数据
    DelData: function (value) {

        $('#_hrId').val(value);

        $('#result1').html('确定要删除吗？此操作无法撤销！');
        $("#myModal1").modal("show");
    },
    //确定删除
    DelDone: function () {

        $("#myModal1").modal("toggle");

        $.ajax({
            type: 'POST',
            url: '../Hrm/DelHr_Bais_xmjz',
            data: { id: $('#_hrId').val() },
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

            window.open("../Hrm/ExportHr_Bais_xmjz?startDate=" + _startDate + "&endDate=" + _endDate);
        }
        else {
            $('#result').html("当前无数据导出！");
            $("#myModal").modal("show");
        }
    },



    //返回
    Back: function () {

        $('#addData').hide();
        $('#import').hide();
        $('#mainData').show();  //主信息

        jsObject.Setted.InitTable();
    }

};

//注册事件
operateEvents = {

    'click .RoleOfB': function (e, value, row, index) {
        jsObject.Setted.DelData(value); //删除
    }
};
