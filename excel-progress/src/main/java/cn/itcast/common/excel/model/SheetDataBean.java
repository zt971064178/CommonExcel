package cn.itcast.common.excel.model;

import java.util.List;

/**
 * 创建多个sheet的时候每个sheet的参数，不支持sheet的分页
 * Created by zhangtian on 2017/4/27.
 */
public class SheetDataBean {
    private List<?> appDatas;
    private Class<?> clazz;
    private String sheetNames;
    private boolean isBigData = Boolean.FALSE; //默认不支持分页
    private int pageSize = 0;

    public List<?> getAppDatas() {
        return appDatas;
    }

    public void setAppDatas(List<?> appDatas) {
        this.appDatas = appDatas;
    }

    public Class<?> getClazz() {
        return clazz;
    }

    public void setClazz(Class<?> clazz) {
        this.clazz = clazz;
    }

    public String getSheetNames() {
        return sheetNames;
    }

    public void setSheetNames(String sheetNames) {
        this.sheetNames = sheetNames;
    }

    public boolean isBigData() {
        return isBigData;
    }

    public void setBigData(boolean bigData) {
        isBigData = bigData;
    }

    public int getPageSize() {
        return pageSize;
    }

    public void setPageSize(int pageSize) {
        this.pageSize = pageSize;
    }
}
