package com.hm.util.excel.base;


import java.io.Serializable;

/**
 * 后台页面返回基础数据
 * Created by Administrator on 2018-5-10.
 */
public class BizResult<T> implements Serializable {

    private static final long serialVersionUID = -4577255781088498763L;
    private int code;  //状态码
    private String msg;  //描述信息
    private T data;  //服务端返回数据

    public int getCode() {
        return code;
    }

    public void setCode(int code) {
        this.code = code;
    }

    public String getMsg() {
        return msg;
    }

    public void setMsg(String msg) {
        this.msg = msg;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    //定义默认
    public BizResult() {}

    /**
     * 自定义返回状态码和描述信息
     */
    public static<T>  BizResult createSuccessResult(String msg,T data){
        BizResult result = new BizResult();
        result.setCode(0);
        result.setMsg(msg);
        result.setData(data);
        return result;
    }
    public static BizResult createFailResult(String msg){
        BizResult result = new BizResult();
        result.setCode(1);
        result.setMsg(msg);
        return result;
    }

    /**
     * 自定义返回所有信息
     */
    public static <T> BizResult createResult(int code,String msg,T data){
        BizResult result = new BizResult();
        result.setCode(code);
        result.setMsg(msg);
        result.setData(data);
        return result;
    }
}
