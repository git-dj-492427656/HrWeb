
var jsObject = {};

jsObject.Setted = {

    init: function () {

        nowYear: '';    //当前年度
        isExp: '0';     //导出判断

        //初始化
        jsObject.Setted.onLoad();

        $("#yearly").change(function () {

            jsObject.Setted.GetData();
        });

        $("#htType").change(function () {

            jsObject.Setted.GetData();
        });

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

    },

    //页面初始化
    onLoad: function () {

        $('#import').hide();    //上传文件
        $('#mainData').show();  //主信息

        //日期选择 初始值
        var myDate = new Date();
        var myYear = myDate.getFullYear();      //当前年度
        jsObject.Setted.nowYear = myYear;

        jsObject.Setted.SetYearly();
    },

    //设置年度
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

                    jsObject.Setted.GetData();
                }
            }
        });
    },

    //获取数据
    GetData: function () {

        $("#colData").empty();
        jsObject.Setted.isExp = "0";

        $.ajax({
            type: 'POST',
            url: '../Hrm/GetHr_Rule_ht',
            data: { yearly: $('#yearly').val(), htType: $('#htType').val() },
            success: function (data) {

                var json = eval("(" + data + ")");

                if (json.isExp == "1") {
                    jsObject.Setted.isExp = json.isExp;
                    $("#colData").html(json.dataStr);
                }
                else {
                    $("#colData").html("无数据！");
                }
            }
        });
    },

    //导出数据
    Export: function () {

        if (jsObject.Setted.isExp == "1") {

            var yearly = $('#yearly').val();
            var htType = $('#htType').val();
            window.open("../Hrm/ExportHr_Rule_ht?yearly=" + yearly + "&htType=" + htType);
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

        jsObject.Setted.GetData();
    }

};
