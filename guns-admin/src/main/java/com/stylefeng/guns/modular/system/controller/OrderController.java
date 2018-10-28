package com.stylefeng.guns.modular.system.controller;

import com.baomidou.mybatisplus.mapper.EntityWrapper;
import com.stylefeng.guns.core.base.controller.BaseController;
import com.stylefeng.guns.core.common.annotion.Permission;
import com.stylefeng.guns.core.common.constant.Const;
import com.stylefeng.guns.core.util.DateUtil;
import com.stylefeng.guns.core.util.ToolUtil;
import com.stylefeng.guns.modular.system.excel.ExcelData;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.*;
import org.springframework.ui.Model;
import org.springframework.beans.factory.annotation.Autowired;
import com.stylefeng.guns.core.log.LogObjectHolder;
import com.stylefeng.guns.modular.system.model.Order;
import com.stylefeng.guns.modular.system.service.IOrderService;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import static com.stylefeng.guns.modular.system.excel.ExportExcelUtils.exportExcel;

/**
 * 订单管理控制器
 *
 * @author fengshuonan
 * @Date 2018-10-20 00:25:00
 */
@Controller
@RequestMapping("/order")
public class OrderController extends BaseController {

    private String PREFIX = "/system/order/";

    @Autowired
    private IOrderService orderService;

    /**
     * 跳转到订单管理首页
     */
    @RequestMapping("")
    public String index() {
        return PREFIX + "order.html";
    }

    /**
     * 跳转到今日未完成订单
     */
    @RequestMapping("/order_unfinished")
    public String unfinishedOrder() {
        return PREFIX + "order_unfinished.html";
    }

    /**
     * 跳转到明日未完成订单
     */
    @RequestMapping("/order_unfinished_tomorrow")
    public String unfinishedOrderTomorrow() {
        return PREFIX + "order_unfinished_tomorrow.html";
    }

    /**
     * 跳转到已完成订单
     */
    @RequestMapping("/order_finished")
    public String finishedOrder() {
        return PREFIX + "order_finished.html";
    }

    /**
     * 跳转到已完成订单
     */
    @RequestMapping("/order_exceptional")
    public String exceptionalOrder() {
        return PREFIX + "order_exceptional.html";
    }

    /**
     * 跳转到添加订单管理
     */
    @RequestMapping("/order_add")
    public String orderAdd() {
        return PREFIX + "order_add.html";
    }

    /**
     * 跳转到修改订单管理
     */
    @RequestMapping("/order_update/{orderId}")
    public String orderUpdate(@PathVariable Integer orderId, Model model) {
        Order order = orderService.selectById(orderId);
        model.addAttribute("item", order);
        LogObjectHolder.me().set(order);
        return PREFIX + "order_edit.html";
    }

    /**
     * 未完成订单列表
     */
    @RequestMapping(value = "/unfinishedList")
    @ResponseBody
    public Object unfinishedList(String telephone, String orderTime, String expressageCompany) {
        String nowTime = DateUtil.getDay();
        String lastTime = DateUtil.getBeforeDate();
        if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().between("order_time", lastTime.concat(" 16:30:01"), nowTime.concat(" 16:30:00"));
//            entityWrapper.eq("expressage_status", 2).andNew().between("order_time", nowTime.concat(" 00:00:00"), nowTime.concat(" 16:30:00")).orNew().between("order_time",lastTime.concat(" 16:30:01"),lastTime.concat(" 23:59:59"));
            System.out.println(entityWrapper.getSqlSegment());
//            entityWrapper.eq("expressage_status", 2);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("customer_telephone", telephone);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("customer_telephone", telephone).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("customer_telephone", telephone).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("expressage_company", expressageCompany).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("customer_telephone", telephone).and().like("order_time", orderTime).and()
                    .like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else {
            return orderService.selectList(null);
        }
//        return orderService.selectList(null);
    }

    /**
     * 已经完成订单列表
     */
    @RequestMapping(value = "/finishedList")
    @ResponseBody
    public Object finishedList(String telephone, String orderTime, String expressageCompany) {
//        if (ToolUtil.isEmpty(telephone)){
//            return orderService.selectList(null);
//        }
//        else {
//            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
//            entityWrapper.like("customer_telephone",telephone);
//            return orderService.selectList(entityWrapper);
//        }
        if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("Id", telephone);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("Id", telephone).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("Id", telephone).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("expressage_company", expressageCompany).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 1).and().like("Id", telephone).and().like("order_time", orderTime).and()
                    .like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else {
            return orderService.selectList(null);
        }
    }

    /**
     * 异常订单列表
     */
    @RequestMapping(value = "/exceptionalList")
    @ResponseBody
    public Object exceptionalList(String telephone, String orderTime, String expressageCompany) {
//        if (ToolUtil.isEmpty(telephone)){
//            return orderService.selectList(null);
//        }
//        else {
//            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
//            entityWrapper.like("customer_telephone",telephone);
//            return orderService.selectList(entityWrapper);
//        }
        if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("customer_telephone", telephone);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("customer_telephone", telephone).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("customer_telephone", telephone).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("expressage_company", expressageCompany).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 3).and().like("customer_telephone", telephone).and().like("order_time", orderTime).and()
                    .like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else {
            return orderService.selectList(null);
        }
    }

    /**
     * 今日待领订单列表
     */
    @RequestMapping(value = "/waitingList")
    @ResponseBody
    public Object waitingList(String telephone, String orderTime, String expressageCompany) {
//        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        String nowTime = DateUtil.getDay();
//        entityWrapper.between("order_time", nowTime.concat(" 00:00:00"), nowTime.concat(" 16:30:00"));
//        return orderService.selectList(entityWrapper);
        if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
//            entityWrapper.eq("expressage_status", 2);
            entityWrapper.eq("expressage_status", 2).and().between("order_time", nowTime.concat(" 16:30:01"), nowTime.concat(" 23:59:59"));
            System.out.println(entityWrapper.getSqlSegment());
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("Id", telephone);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("Id", telephone).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("Id", telephone).and().like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("expressage_company", expressageCompany).and().like("order_time", orderTime);
            return orderService.selectList(entityWrapper);
        } else if (ToolUtil.isNotEmpty(telephone) && ToolUtil.isNotEmpty(orderTime) && ToolUtil.isNotEmpty(expressageCompany)) {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.eq("expressage_status", 2).and().like("Id", telephone).and().like("order_time", orderTime).and()
                    .like("expressage_company", expressageCompany);
            return orderService.selectList(entityWrapper);
        } else {
            return orderService.selectList(null);
        }
    }

    @RequestMapping(value = "/list")
    @ResponseBody
    public Object list(String telephone, String orderTime, String expressageCompany) {
        if (ToolUtil.isEmpty(telephone)) {
            return orderService.selectList(null);
        } else {
            EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
            entityWrapper.like("customer_telephone", telephone);
            return orderService.selectList(entityWrapper);
        }

    }

    /**
     * 新增订单管理
     */
    @RequestMapping(value = "/add")
//    @Permission(Const.ADMIN_NAME)
    @ResponseBody
    public Object add(Order order) {
        orderService.insert(order);
        return SUCCESS_TIP;

    }

    @GetMapping("/wxadd")
    @ResponseBody
    public Object wxadd(){
        return SUCCESS_TIP;
    }

    /**
     * 删除订单管理
     */
    @RequestMapping(value = "/delete")
//    @Permission(Const.ADMIN_NAME)
    @ResponseBody
    public Object delete(@RequestParam Integer orderId){
        orderService.deleteById(orderId);
        return SUCCESS_TIP;
    }

    /**
     * 修改订单管理
     */
    @RequestMapping(value = "/update")
//    @Permission(Const.ADMIN_NAME)
    @ResponseBody
    public Object update(Order order) {
        orderService.updateById(order);
        System.out.println(orderService.updateById(order));
        return SUCCESS_TIP;
    }

    /**
     * 订单管理详情
     */
    @RequestMapping(value = "/detail/{orderId}")
    @ResponseBody
    public Object detail(@PathVariable("orderId") Integer orderId) {
        return orderService.selectById(orderId);
    }

    /**
     * 一键完成订单
     */
    @RequestMapping("/completeOrderToday")
    @ResponseBody
    public Object completeOrder() {
        String nowTime = DateUtil.getDay();
        String lastTime = DateUtil.getBeforeDate();
        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        entityWrapper.eq("expressage_status", 2).and().between("order_time", lastTime.concat(" 16:30:01"), nowTime.concat(" 16:30:00"));
        Order order = new Order();
        order.setExpressageStatus(1);
        orderService.update(order, entityWrapper);
        return SUCCESS_TIP;
    }

    /**
     * 一键完成明日订单
     */
    @RequestMapping("/completeOrderTomorrow")
    @ResponseBody
    public Object completeOrderTomorrow() {
        String nowTime = DateUtil.getDay();
        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        entityWrapper.eq("expressage_status", 2).and().between("order_time", nowTime.concat(" 16:30:01"), nowTime.concat(" 23:59:59"));;
        Order order = new Order();
        order.setExpressageStatus(1);
        orderService.update(order, entityWrapper);
        return SUCCESS_TIP;
    }

//    @RequestMapping("/test1")
//    @ResponseBody
//    public void test1(HttpServletResponse response) throws Exception {
////        response.setContentType("application/vnd..ms-excel");
////        response.setHeader("content-Disposition","attachment;filename="+ URLEncoder.encode("order.xlsx","utf-8"));
//        ExcelData data = new ExcelData();
//        data.setName("订单");
//        List<String> titles = new ArrayList();
//        titles.add("订单编号");
//        titles.add("用户");
//        titles.add("收件人");
//        titles.add("手机号码");
//        titles.add("收货地址");
//        titles.add("快递公司");
//        titles.add("取件码");
//        titles.add("到件日期");
//        titles.add("订单状态");
//        titles.add("下单时间");
//        titles.add("备注");
//        data.setTitles(titles);
//
//        List<List<Object>> rows = new ArrayList();
//        List<Object> row = new ArrayList();
//        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
//        entityWrapper.eq("expressage_status",2);
//        List<Order> orderList = orderService.selectList(entityWrapper);
//        for (int i =0;i<orderList.size();i++){
//            row.add(orderList.get(i).getId());
//            row.add(orderList.get(i).getCustomerName());
//            row.add(orderList.get(i).getRecipients());
//            row.add(orderList.get(i).getCustomerTelephone());
//            row.add(orderList.get(i).getCustomerPlace());
//            row.add(orderList.get(i).getExpressageCompany());
//            row.add(orderList.get(i).getExpressageCode());
//            row.add(orderList.get(i).getExpressageArriveTime());
//            row.add(orderList.get(i).getExpressageStatus());
//            row.add(orderList.get(i).getOrderTime());
//            row.add(orderList.get(i).getExpressageDescription());
//            rows.add(row);
//        }
//
//
//        data.setRows(rows);
//
////        rows.add(row);
////        data.setExrow(11);
////        for (int i=0;i<rows.size();i++){
////            data.setRows(rows.subList(i,i+1));
////
////        }
//        //生成本地
//        /*File f = new File("c:/test.xlsx");
//        FileOutputStream out = new FileOutputStream(f);
//        ExportExcelUtils.exportExcel(data, out);
//        out.close();*/
//        ExportExcelUtils.exportExcel(response, "order1.xlsx", data);
////        workbook.write(response.getOutputStream());
////        orderService.deleteById(orderId);
////        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
////        entityWrapper.eq("expressage_status",2);
////        Order order = new Order();
////        order.setExpressageStatus(1);
////        orderService.update(order,entityWrapper);
////        return null;
////        return SUCCESS_TIP;
//    }


    /*
     * 导出今日已完成订单
     */
    @RequestMapping(value = "/exportFinished")
    @ResponseBody
    public void exportFinished(HttpServletResponse response) throws IOException {
        String nowTime = DateUtil.getDay();
        String lastTime = DateUtil.getBeforeDate();
        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        entityWrapper.eq("expressage_status",1).and().between("order_time", lastTime.concat(" 16:30:01"), nowTime.concat(" 16:30:00"));
        List<Order> users = orderService.selectList(entityWrapper);

        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet("配送订单");

        HSSFRow row = null;

//        row = sheet.createRow(0);//创建第一个单元格
//        row.setHeight((short) (40 * 25));
//        row.createCell(0).setCellValue("今日配送订单列表");//为第一行单元格设值

        /*为标题设计空间
         * firstRow从第1行开始
         * lastRow从第0行结束
         *
         *从第1个单元格开始
         * 从第3个单元格结束
         */
//        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 2);
//        sheet.addMergedRegion(rowRegion);

      /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
      sheet.addMergedRegion(columnRegion);*/


        /*
         * 动态获取数据库列 sql语句 select COLUMN_NAME from INFORMATION_SCHEMA.Columns where table_name='user' and table_schema='test'
         * 第一个table_name 表名字
         * 第二个table_name 数据库名称
         * */
        row = sheet.createRow(0);
//        row.setHeight((short) (60 * 40));//设置行高
        row.createCell(0).setCellValue("订单编号");//为第一个单元格设值
        row.createCell(1).setCellValue("用户");//为第二个单元格设值
        row.createCell(2).setCellValue("收件人");//为第三个单元格设值
        row.createCell(3).setCellValue("手机号码");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(4).setCellValue("收货地址");//为第三个单元格设值
        row.createCell(5).setCellValue("快递公司");//为第三个单元格设值
        row.createCell(6).setCellValue("取件码");//为第三个单元格设值
        row.createCell(7).setCellValue("到件日期");//为第三个单元格设值
        row.createCell(8).setCellValue("订单状态");//为第三个单元格设值
        row.createCell(9).setCellValue("下单时间");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(10).setCellValue("备注");//为第三个单元格设值

        for (int i = 0; i < users.size(); i++) {
            sheet.setDefaultRowHeight((short) (20 * 40));
            row.setHeight((short) (30 * 20));//设置行高
            row = sheet.createRow(i + 2);
            Order user = users.get(i);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getCustomerName());
            row.createCell(2).setCellValue(user.getRecipients());
            row.createCell(3).setCellValue(user.getCustomerTelephone());
            row.createCell(4).setCellValue(user.getCustomerPlace());
            row.createCell(5).setCellValue(user.getExpressageCompany());
            row.createCell(6).setCellValue(user.getExpressageCode());
            row.createCell(7).setCellValue(user.getExpressageArriveTime());
            row.createCell(8).setCellValue(user.getExpressageStatus());
            row.createCell(9).setCellValue(user.getOrderTime());
            row.createCell(10).setCellValue(user.getExpressageDescription());
        }
//        sheet.setDefaultRowHeight((short) (20 * 30));
        //列宽自适应
        for (int i = 0; i <= 13; i++) {
            sheet.autoSizeColumn(i);
        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        response.setHeader("Content-disposition", "attachment;filename=order.xls");//默认Excel名称
        wb.write(os);
        os.flush();
        os.close();


    }

    /*
     * 导出今日未完成订单
     */
    @RequestMapping(value = "/exportUnFinishedToworrow")
    @ResponseBody
    public void exportUnFinishedToworrow(HttpServletResponse response) throws IOException {
        String nowTime = DateUtil.getDay();
        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        entityWrapper.eq("expressage_status",2).and().between("order_time", nowTime.concat(" 16:30:01"), nowTime.concat(" 23:59:59"));;
        List<Order> users = orderService.selectList(entityWrapper);

        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet("配送订单");

        HSSFRow row = null;

//        row = sheet.createRow(0);//创建第一个单元格
//        row.setHeight((short) (40 * 25));
//        row.createCell(0).setCellValue("今日配送订单列表");//为第一行单元格设值

        /*为标题设计空间
         * firstRow从第1行开始
         * lastRow从第0行结束
         *
         *从第1个单元格开始
         * 从第3个单元格结束
         */
//        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 2);
//        sheet.addMergedRegion(rowRegion);

      /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
      sheet.addMergedRegion(columnRegion);*/


        /*
         * 动态获取数据库列 sql语句 select COLUMN_NAME from INFORMATION_SCHEMA.Columns where table_name='user' and table_schema='test'
         * 第一个table_name 表名字
         * 第二个table_name 数据库名称
         * */
        row = sheet.createRow(0);
//        row.setHeight((short) (60 * 40));//设置行高
        row.createCell(0).setCellValue("订单编号");//为第一个单元格设值
        row.createCell(1).setCellValue("用户");//为第二个单元格设值
        row.createCell(2).setCellValue("收件人");//为第三个单元格设值
        row.createCell(3).setCellValue("手机号码");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(4).setCellValue("收货地址");//为第三个单元格设值
        row.createCell(5).setCellValue("快递公司");//为第三个单元格设值
        row.createCell(6).setCellValue("取件码");//为第三个单元格设值
        row.createCell(7).setCellValue("到件日期");//为第三个单元格设值
        row.createCell(8).setCellValue("订单状态");//为第三个单元格设值
        row.createCell(9).setCellValue("下单时间");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(10).setCellValue("备注");//为第三个单元格设值

        for (int i = 0; i < users.size(); i++) {
            sheet.setDefaultRowHeight((short) (20 * 40));
            row.setHeight((short) (30 * 20));//设置行高
            row = sheet.createRow(i + 2);
            Order user = users.get(i);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getCustomerName());
            row.createCell(2).setCellValue(user.getRecipients());
            row.createCell(3).setCellValue(user.getCustomerTelephone());
            row.createCell(4).setCellValue(user.getCustomerPlace());
            row.createCell(5).setCellValue(user.getExpressageCompany());
            row.createCell(6).setCellValue(user.getExpressageCode());
            row.createCell(7).setCellValue(user.getExpressageArriveTime());
            row.createCell(8).setCellValue(user.getExpressageStatus());
            row.createCell(9).setCellValue(user.getOrderTime());
            row.createCell(10).setCellValue(user.getExpressageDescription());
        }
//        sheet.setDefaultRowHeight((short) (20 * 30));
        //列宽自适应
        for (int i = 0; i <= 13; i++) {
            sheet.autoSizeColumn(i);
        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        response.setHeader("Content-disposition", "attachment;filename=order.xls");//默认Excel名称
        wb.write(os);
        os.flush();
        os.close();


    }

    /*
     * 导出今日未完成订单
     */
    @RequestMapping(value = "/exportUnFinishedToday")
    @ResponseBody
    public void exportUnFinishedToday(HttpServletResponse response) throws IOException {
        String nowTime = DateUtil.getDay();
        String lastTime = DateUtil.getBeforeDate();
        EntityWrapper<Order> entityWrapper = new EntityWrapper<>();
        entityWrapper.eq("expressage_status",2).and().between("order_time", lastTime.concat(" 16:30:01"), nowTime.concat(" 16:30:00"));
        List<Order> users = orderService.selectList(entityWrapper);

        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet("配送订单");

        HSSFRow row = null;

//        row = sheet.createRow(0);//创建第一个单元格
//        row.setHeight((short) (40 * 25));
//        row.createCell(0).setCellValue("今日配送订单列表");//为第一行单元格设值

        /*为标题设计空间
         * firstRow从第1行开始
         * lastRow从第0行结束
         *
         *从第1个单元格开始
         * 从第3个单元格结束
         */
//        CellRangeAddress rowRegion = new CellRangeAddress(0, 0, 0, 2);
//        sheet.addMergedRegion(rowRegion);

      /*CellRangeAddress columnRegion = new CellRangeAddress(1,4,0,0);
      sheet.addMergedRegion(columnRegion);*/


        /*
         * 动态获取数据库列 sql语句 select COLUMN_NAME from INFORMATION_SCHEMA.Columns where table_name='user' and table_schema='test'
         * 第一个table_name 表名字
         * 第二个table_name 数据库名称
         * */
        row = sheet.createRow(0);
//        row.setHeight((short) (60 * 40));//设置行高
        row.createCell(0).setCellValue("订单编号");//为第一个单元格设值
        row.createCell(1).setCellValue("用户");//为第二个单元格设值
        row.createCell(2).setCellValue("收件人");//为第三个单元格设值
        row.createCell(3).setCellValue("手机号码");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(4).setCellValue("收货地址");//为第三个单元格设值
        row.createCell(5).setCellValue("快递公司");//为第三个单元格设值
        row.createCell(6).setCellValue("取件码");//为第三个单元格设值
        row.createCell(7).setCellValue("到件日期");//为第三个单元格设值
        row.createCell(8).setCellValue("订单状态");//为第三个单元格设值
        row.createCell(9).setCellValue("下单时间");//为第三个单元格设值
//        sheet.setDefaultRowHeight((short) (40 * 60));
        row.createCell(10).setCellValue("备注");//为第三个单元格设值

        for (int i = 0; i < users.size(); i++) {
            sheet.setDefaultRowHeight((short) (20 * 40));
            row.setHeight((short) (30 * 20));//设置行高
            row = sheet.createRow(i + 2);
            Order user = users.get(i);
            row.createCell(0).setCellValue(user.getId());
            row.createCell(1).setCellValue(user.getCustomerName());
            row.createCell(2).setCellValue(user.getRecipients());
            row.createCell(3).setCellValue(user.getCustomerTelephone());
            row.createCell(4).setCellValue(user.getCustomerPlace());
            row.createCell(5).setCellValue(user.getExpressageCompany());
            row.createCell(6).setCellValue(user.getExpressageCode());
            row.createCell(7).setCellValue(user.getExpressageArriveTime());
            row.createCell(8).setCellValue(user.getExpressageStatus());
            row.createCell(9).setCellValue(user.getOrderTime());
            row.createCell(10).setCellValue(user.getExpressageDescription());
        }
//        sheet.setDefaultRowHeight((short) (20 * 30));
        //列宽自适应
        for (int i = 0; i <= 13; i++) {
            sheet.autoSizeColumn(i);
        }

        response.setContentType("application/vnd.ms-excel;charset=utf-8");
        OutputStream os = response.getOutputStream();
        response.setHeader("Content-disposition", "attachment;filename=order.xls");//默认Excel名称
        wb.write(os);
        os.flush();
        os.close();

    }
}
