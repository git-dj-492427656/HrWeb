
var jsObject = {};

var dataJson; //表格数据json

jsObject.Setted = {

    init: function () {

        //初始化
        jsObject.Setted.onLoad();

        //导入数据
        $('#btnImp').click(function () {

            $('#mainData').hide();
            $('#import').show();

        });

        $('#btnBack2').click(function () {
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

        $('#import').hide();    //上传文件
        $('#mainData').show();  //主信息

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
                    field: 'mType',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '120',
                    sortable: true,
                    title: '类  别'
                },
                {
                    field: 'ruleCode',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '120',
                    sortable: true,
                    title: '岗位配备编码'
                },
                {
                    field: 'ruleName',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '120',
                    sortable: true,
                    title: '岗位配备名称'
                },
                {
                    field: 'easCode',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '120',
                    sortable: true,
                    title: 'eas编码'
                },
                {
                    field: 'easName',
                    align: 'center',
                    halign: 'center',
                    valign: 'middle',
                    width: '120',
                    sortable: true,
                    title: 'eas名称'
                }
            ],

            url: '../Hrm/GetHr_Middle_rule',        //后台数据url

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
            url: '../Hrm/DelHr_Middle_rule',
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
            
            window.open("../Hrm/ExportHr_Middle_rule");
        }
        else {
            $('#result').html("当前无数据导出！");
            $("#myModal").modal("show");
        }
    },



    //返回
    Back: function () {

        $('#import').hide();
        $('#mainData').show();  //主信息

        jsObject.Setted.InitTable();
    }

};

