
var jsObject = {};

//树
var setting;

//当前节点
var _node;

jsObject.Setted = {

    init: function () {

        //初始化
        $('#search').css('display', 'none');

        setting = {
            view: {
                dblClickExpand: false, //双击展开
                selectedMulti: false, //同时选中多个
                fontCss: { 'color': '#515356', "font-weight": "bold" }, //字体样式
                showIcon: false,
                showLine: true,
                showTitle: false
            },
            data: {
                simpleData: {
                    enable: true,
                    idKey: "id",
                    pIdKey: "pId",
                    rootPId: ""
                },
                key: {
                    type: "type", //自定义字段
                    memo: "memo"  //自定义字段
                }
            },
            callback: {
                onClick: jsObject.Setted.zTreeOnClick  //单击
            },
            check: {
                enable: true
            }
        }

        jsObject.Setted.GetComList();

        $('#btnOk').click(function () {
            jsObject.Setted.ChangeCom();
        });

        $('#btnSearch').click(function () {
            
            jsObject.Setted.GetComList();
        });

        $('#btnClose').click(function () {
            $('#myModal').modal('toggle');
            window.location = "../Main/Index";
        });
    },

    GetComList: function () {
        $.ajax({
            type: 'POST',
            url: '../Main/GetComList',
            data: { comName: $('#comName1').val() },
            success: function (data) {
                var nodes = eval('(' + data + ')');

                $.fn.zTree.init($("#tree"), setting, nodes);
            }
        });
    },

    //单击
    zTreeOnClick: function (event, treeId, treeNode) {

        _node = treeNode;

        //判断是否有子节点
        if (treeNode.isParent) {
            var zTree = $.fn.zTree.getZTreeObj("tree");
            zTree.expandNode(treeNode); //展开
        }

        $('#comName2').val(treeNode.name);
    },

    // 获取子节点集合
    getAllChildrenNodes: function (treeNode, result) {
        if (treeNode.isParent) {
            var childrenNodes = treeNode.children;
            if (childrenNodes) {
                for (var i = 0; i < childrenNodes.length; i++) {
                    result += ',' + childrenNodes[i].id;
                    result = jsObject.Setted.getAllChildrenNodes(childrenNodes[i], result);
                }
            }
        }
        return result;
    },

    // 确认切换
    ChangeCom: function () {

        if ($('#comName2').val() == "") {
            return;
        }

        //节点集合
        var _ids = _node.id;

        //判断是否有子节点
        if (_node.isParent) {
            var result = _node.id;
            _ids = jsObject.Setted.getAllChildrenNodes(_node, result);
        }

        //更改后台数据
        $.ajax({
            type: 'POST',
            url: '../Main/ChangeCom',
            data: {
                ids: _ids, comId: _node.id, comName: _node.name, comNumber: _node.fatherId, fatherId: _node.number,
                isSub: _node.isSub, EASnumber: _node.EASnumber, EASkcnumber: _node.EASkcnumber, property: _node.property,
                gsmid: _node.gsmid
            },
            success: function (data) {

                $('#result').html("已进入管理单元：" + _node.name);
                $('#myModal').modal('show');

            }
        });
    }



};
