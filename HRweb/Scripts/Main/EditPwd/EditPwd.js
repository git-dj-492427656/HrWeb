
var jsObject = {};

jsObject.Setted = {

    init: function () {

        // 初始化
        jsObject.Setted.ClearPwd();

        //保存
        $('#btnSave').click(function () {
            jsObject.Setted.SavaPwd();
        });

        //清空
        $('#btnClear').click(function () {
            jsObject.Setted.ClearPwd();
        });

    },

    SavaPwd: function () {

        $.ajax({
            type: 'POST',
            url: '../Main/SavePwd',
            data: { oldPwd: $('#oldPwd').val(), newPwd: $('#newPwd').val(), newPwd2: $('#newPwd2').val() },
            success: function (data) {

                $('#result').html(data);

                $('#myModal').modal('show');
            }
        });
    },

    ClearPwd: function () {

        $('#form2')[0].reset();
    }

};




