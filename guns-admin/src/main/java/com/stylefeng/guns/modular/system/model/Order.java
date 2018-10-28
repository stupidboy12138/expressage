package com.stylefeng.guns.modular.system.model;

import com.baomidou.mybatisplus.enums.IdType;
import com.baomidou.mybatisplus.annotations.TableId;
import com.baomidou.mybatisplus.annotations.TableField;
import com.baomidou.mybatisplus.activerecord.Model;
import com.baomidou.mybatisplus.annotations.TableName;

import java.io.Serializable;
import java.sql.Date;

/**
 * <p>
 * 
 * </p>
 *
 * @author AlbertXiao
 * @since 2018-10-20
 */
@TableName("sys_order")
public class Order extends Model<Order> {

    private static final long serialVersionUID = 1L;

    @TableId(value = "Id", type = IdType.AUTO)
    private Integer Id;
    @TableField("customer_name")
    private String customerName;
    @TableField("customer_telephone")
    private String customerTelephone;
    @TableField("expressage_company")
    private String expressageCompany;
    @TableField("customer_place")
    private String customerPlace;
    @TableField("expressage_arrive_time")
    private String expressageArriveTime;
    @TableField("expressage_description")
    private String expressageDescription;
    @TableField("expressage_code")
    private String expressageCode;
    @TableField("expressage_status")
    private Integer expressageStatus;
    @TableField("recipients")
    private String recipients;
    @TableField ("order_time")
    private Date orderTime;
    @TableField("specification")
    private int specification;



    @Override
    protected Serializable pkVal() {
        return this.Id;
    }

    public static long getSerialVersionUID() {
        return serialVersionUID;
    }

    public Integer getId() {
        return Id;
    }

    public void setId(Integer id) {
        Id = id;
    }

    public String getCustomerName() {
        return customerName;
    }

    public void setCustomerName(String customerName) {
        this.customerName = customerName;
    }

    public String getCustomerTelephone() {
        return customerTelephone;
    }

    public void setCustomerTelephone(String customerTelephone) {
        this.customerTelephone = customerTelephone;
    }

    public String getExpressageCompany() {
        return expressageCompany;
    }

    public void setExpressageCompany(String expressageCompany) {
        this.expressageCompany = expressageCompany;
    }

    public String getCustomerPlace() {
        return customerPlace;
    }

    public void setCustomerPlace(String customerPlace) {
        this.customerPlace = customerPlace;
    }

    public String getExpressageArriveTime() {
        return expressageArriveTime;
    }

    public void setExpressageArriveTime(String expressageArriveTime) {
        this.expressageArriveTime = expressageArriveTime;
    }

    public String getExpressageDescription() {
        return expressageDescription;
    }

    public void setExpressageDescription(String expressageDescription) {
        this.expressageDescription = expressageDescription;
    }

    public String getExpressageCode() {
        return expressageCode;
    }

    public void setExpressageCode(String expressageCode) {
        this.expressageCode = expressageCode;
    }

    public Integer getExpressageStatus() {
        return expressageStatus;
    }

    public void setExpressageStatus(Integer expressageStatus) {
        this.expressageStatus = expressageStatus;
    }

    public String getRecipients() {
        return recipients;
    }

    public void setRecipients(String recipients) {
        this.recipients = recipients;
    }

    public Date getOrderTime() {
        return orderTime;
    }

    public void setOrderTime(Date orderTime) {
        this.orderTime = orderTime;
    }

    public int getSpecification() {
        return specification;
    }

    public void setSpecification(int specification) {
        this.specification = specification;
    }

    @Override
    public String toString() {
        return "Order{" +
                "Id=" + Id +
                ", customerName='" + customerName + '\'' +
                ", customerTelephone='" + customerTelephone + '\'' +
                ", expressageCompany='" + expressageCompany + '\'' +
                ", customerPlace='" + customerPlace + '\'' +
                ", expressageArriveTime='" + expressageArriveTime + '\'' +
                ", expressageDescription='" + expressageDescription + '\'' +
                ", expressageCode='" + expressageCode + '\'' +
                ", expressageStatus=" + expressageStatus +
                ", recipients='" + recipients + '\'' +
                ", orderTime=" + orderTime +
                ", specification=" + specification +
                '}';
    }
}
