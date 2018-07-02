
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
        $('#selDate').change(function () {

            //jsObject.Setted.InitTable();
        });


        //导出数据
        $('#btnChart').click(function () {

            $('#result').html("未更新..");
            $("#myModal").modal("show");
        });

        //导出数据
        $('#btnExp').click(function () {

            jsObject.Setted.Export();
        });

    },

    //页面初始化
    onLoad: function () {

        //日期选择 初始值
        var myDate = new Date();
        var myYear = myDate.getFullYear();      //当前年度
        var myMonth = myDate.getMonth() + 1;    //当前月度
        
        //初始查询条件
        jsObject.Setted.nowYear = myYear;
        jsObject.Setted.SetYearly(); 


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
            uniqueId: "id",                     //每一行的唯一标识，一般为主键列
            showToggle: true,                    //是否显示详细视图和列表视图的切换按钮
            cardView: false,                    //是否显示详细视图
            detailView: false,                   //是否显示父子表

            //表格列定义
            columns: [
                {
                    field: 'id',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '序号'
                },
                {
                    field: 'classify',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '项目/分类'
                },
                {
                    field: 'goal',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '年度目标'
                },
                {
                    field: 'month1',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '1月'
                },
                {
                    field: 'month2',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '2月'
                },
                {
                    field: 'month3',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '3月'
                },
                {
                    field: 'month4',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '4月'
                },
                {
                    field: 'month5',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '5月'
                },
                {
                    field: 'month6',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '6月'
                },
                {
                    field: 'month7',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '7月'
                },
                {
                    field: 'month8',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '8月'
                },
                {
                    field: 'month9',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '9月',
                },
                {
                    field: 'month10',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '10月',
                },
                {
                    field: 'month11',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '11月',
                },
                {
                    field: 'month12',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '100',
                    title: '12月',
                }
            ],

            url: '../Hrm/GetHr_Model',        //后台数据url

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

                return params;
            }
        });
    },

    //导出数据
    Export: function () {

        if (dataJson.length > 0) {

            //var _selDate = $('#selDate').val();

            //window.open("../Hrm/ExportHr_Midrgzc?selDate=" + _selDate);
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

                    jsObject.Setted.InitTable(); 
                }

            }
        });
    }


};
