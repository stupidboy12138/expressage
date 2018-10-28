/**
 * 订单管理管理初始化
 */
var Order = {
    id: "OrderTable1",	//表格id
    seItem: null,		//选中的条目
    table: null,
    layerIndex: -1
};

/**
 * 初始化表格的列
 */
Order.initColumn = function () {
    return [
        {field: 'selectItem', radio: true},
            {title: '订单编号1', field: 'id', visible: true, align: 'center', valign: 'middle'},
            {title: '用户1', field: 'customerName', visible: true, align: 'center', valign: 'middle'},
            {title: '收件人1', field: 'recipients', visible: true, align: 'center', valign: 'middle'},
            {title: '手机号码1', field: 'customerTelephone', visible: true, align: 'center', valign: 'middle'},
            {title: '收货地址1', field: 'customerPlace', visible: true, align: 'center', valign: 'middle'},
            {title: '快递公司1', field: 'expressageCompany', visible: true, align: 'center', valign: 'middle'},
            {title: '取件码', field: 'expressageCode', visible: true, align: 'center', valign: 'middle'},
            {title: '到件日期', field: 'expressageArriveTime', visible: true, align: 'center', valign: 'middle'},
            {title: '订单状态', field: 'expressageStatus', visible: true, align: 'center', valign: 'middle'},
            {title: '下单时间', field: 'orderTime', visible: true, align: 'center', valign: 'middle'},
            {title: '备注', field: 'expressageDescription', visible: true, align: 'center', valign: 'middle'}
    ];
};

/**
 * 检查是否选中
 */
Order.check = function () {
    var selected = $('#' + this.id).bootstrapTable('getSelections');
    if(selected.length == 0){
        Feng.info("请先选中表格中的某一记录！");
        return false;
    }else{
        Order.seItem = selected[0];
        return true;
    }
};

/**
 * 点击添加订单管理
 */
Order.openAddOrder = function () {
    var index = layer.open({
        type: 2,
        title: '添加订单管理',
        area: ['800px', '420px'], //宽高
        fix: false, //不固定
        maxmin: true,
        content: Feng.ctxPath + '/order/order_add'
    });
    this.layerIndex = index;
};

/**
 * 打开查看订单管理详情
 */
Order.openOrderDetail = function () {
    if (this.check()) {
        var index = layer.open({
            type: 2,
            title: '订单管理详情',
            area: ['800px', '420px'], //宽高
            fix: false, //不固定
            maxmin: true,
            content: Feng.ctxPath + '/order/order_update/' + Order.seItem.id
        });
        this.layerIndex = index;
    }
};

/**
 * 删除订单管理
 */
Order.delete = function () {
    if (this.check()) {
        var ajax = new $ax(Feng.ctxPath + "/order/delete", function (data) {
            Feng.success("删除成功!");
            Order.table.refresh();
        }, function (data) {
            Feng.error("删除失败!" + data.responseJSON.message + "!");
        });
        ajax.set("orderId",this.seItem.id);
        ajax.start();
    }
};
/**
 * 导出excel
 */
Order.completeExcel = function(){
    location.href='http://localhost:8888/order/exportUnFinishedToday';
    var ajax = new $ax(Feng.ctxPath + "/order/exportUnFinishedToday", function (data) {
        Feng.success("一键完成订单成功！");
        // Order.table.refresh();
    }, function (data) {
        Feng.error("一键完成订单失败!" + data.responseJSON.message + "!");
    });
    // ajax.set("orderId",this.seItem.id);
    ajax.start();
};

/**
 * 一键完成订单
 */
Order.completeOrder = function(){
    var ajax = new $ax(Feng.ctxPath + "/order/completeOrderToday", function (data) {
        Feng.success("一键完成订单成功！");
        // Order.table.refresh();
    }, function (data) {
        Feng.error("一键完成订单失败!" + data.responseJSON.message + "!");
    });
    // ajax.set("orderId",this.seItem.id);
    ajax.start();
};

/**
 * 查询订单管理列表
 */
Order.search = function () {
    var queryData = {};
    queryData['telephone'] = $("#telephone").val();
    queryData['orderTime'] = $("#orderTime").val();
    queryData['expressageCompany'] = $("#expressageCompany").val();
    Order.table.refresh({query: queryData});
};

$(function () {
    var defaultColunms = Order.initColumn();
    var table = new BSTable(Order.id, "/order/unfinishedList", defaultColunms);
    table.setPaginationType("client");
    Order.table = table.init();
});

