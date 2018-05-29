package com.hm.controller;

import com.hm.pojo.FileVo;
import com.hm.util.excel.ExcelReaderUtil;
import com.hm.util.excel.base.BizResult;
import com.hm.util.excel.common.UploadUtils;
import org.springframework.stereotype.Controller;
import org.springframework.util.CollectionUtils;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;

import javax.servlet.http.HttpServletRequest;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Administrator on 2018/5/29 0029.
 */
@Controller
public class workController {
    private final static Integer BUSINESS_DETAILS = 0;
    private final static Integer TENDERING_MATERIAL = 1;
    private final static String FILE_TEMPLATE_NAME = "成交明细";
    private final static String FILE_FLAG_NAME = "包号";
    private final static String KEY_ONE = "包号";
    private final static String KEY_TWO = "应答人名称";

    @RequestMapping("/index")
    public String index() {
        return "index";
    }

    @RequestMapping("/upload")
    @ResponseBody
    public BizResult upload(String fileName, HttpServletRequest request) {
        UploadUtils up = new UploadUtils();
        BizResult result = up.upload(fileName, request);
        return result;
    }

    @RequestMapping("/excel")
    public void excelRead(FileVo fileVo) {
        //综合排序
        String comprehensive = fileVo.getComprehensive();
        //报名
        String signUp = fileVo.getSignUp();
        //初评汇报
        String PreliminaryReviewReport = fileVo.getPreliminaryReviewReport();
        //采购需求
        String purchasingDemand = fileVo.getPurchasingDemand();
        Integer type = fileVo.getType();
/*        try {
            List<Map<Integer, Object>> comprehensives = this.readExcel(comprehensive,0);
            List<Map<Integer, Object>> signUps = this.readExcel(signUp,0);
            List<Map<Integer, Object>> PreliminaryReviewReports = this.readExcel(PreliminaryReviewReport,0);
            List<Map<Integer, Object>> purchasingDemands = this.readExcel(purchasingDemand,0);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }*/
        if (type == BUSINESS_DETAILS) {
            this.businessDetails(comprehensive, signUp);
        } else if (type == TENDERING_MATERIAL) {
            this.tenderingMaterial();
        }
    }

    public List<Map<Integer, Object>> readExcel(String file, int sheet) throws FileNotFoundException {
        List<Map<Integer, Object>> list = new ArrayList<>();
        if (StringUtils.isEmpty(file)) {
            return list;
        }
        ExcelReaderUtil excelReaderUtil = new ExcelReaderUtil();
        list = excelReaderUtil.readExcelContent(new FileInputStream(file), true, sheet);
        return list;
    }

    //成交明细
    public void businessDetails(String comprehensive, String signUp) {
        try {
            //组装最后的写出List
            List<Map<String, Object>> writeList = new ArrayList<>();

            List<Map<Integer, Object>> signUps = this.readExcel(signUp, 0);
            //获取模板
            UploadUtils up = new UploadUtils();
            List<Map<Integer, Object>> temps = this.readExcel(up.getPath(FILE_TEMPLATE_NAME), 0);
            Map<Integer, Object> tempsMap = temps.get(0);
            //模板字段
            List<String> keys = new ArrayList<>();
            for (int i = 0; true; i++) {
                String key = (String) tempsMap.get(i);
                if (StringUtils.isEmpty(key)) {
                    break;
                } else {
                    keys.add(key);
                }
            }
            for (int i = 0; true; i++) {
                //设置是否定点标志
                Boolean flag = false;
                //动态获取sheet
                List<Map<Integer, Object>> comprehensivesNew = this.readExcel(comprehensive, i);
                //不存在数据则结束
                if (CollectionUtils.isEmpty(comprehensivesNew)) {
                    break;
                } else {
                    //遍历sheet内容
                    for (int j = 0; j < comprehensivesNew.size(); j++) {
                        //获取每行内容
                        Map<Integer, Object> comprehensiveMap = comprehensivesNew.get(j);
                        //获取每行第一个字段 看是不是定点
                        String str = (String) comprehensiveMap.get(0);
                        if (str.indexOf("定点") != -1) {
                            flag = true;
                        }
                        //判断是不是包号
                        if (FILE_FLAG_NAME.equals(str)) {
                            //获取综合排序表主键位置
                            //综合排序表默认主键 0 2
                            Integer comprehensiveMapkeyOne = 0;
                            Integer comprehensiveMapkeyTwo = 2;
                            for (int k = 0; k < comprehensiveMap.size(); k++) {
                                String comprehensiveMapKey = (String) comprehensiveMap.get(k);
                                if (KEY_ONE.equals(comprehensiveMapKey)) {
                                    comprehensiveMapkeyOne=k;
                                }
                                if (KEY_TWO.equals(comprehensiveMapKey)) {
                                    comprehensiveMapkeyTwo=k;
                                }
                            }
                            //获取所需数据 定点取多条 不是定点取一条
                            if (flag) {

                            } else {
                                //取所需数据 如果存在模板字段则取出
                                for (int k = 0; k < comprehensiveMap.size(); k++) {
                                    String comprehensiveMapKey = (String) comprehensiveMap.get(k);
                                    //获取模板字段 判断存在模板字段 取下一行内容
                                    if (keys.contains(comprehensiveMapKey)) {
                                        Map<Integer, Object> comprehensiveAdd = comprehensivesNew.get(j + 1);
                                        String comprehensiveAddValue = (String) comprehensiveAdd.get(k);
                                        Map<String, Object> writeMap = new HashMap<>();
                                        //组装输出map
                                        writeMap.put(comprehensiveMapKey, comprehensiveAddValue);
                                        //补充报名信息
                                        //获取报名信息表头
                                        Map<Integer, Object> signUpsMap = signUps.get(0);
                                        //报名表默认主键 0 4
                                        Integer keyOne = 0;
                                        Integer keyTwo = 4;
                                        for (int l = 0; l < signUpsMap.size(); l++) {
                                            //主键所在位置
                                            String signUpsMapaStr = (String) signUpsMap.get(l);
                                            if (KEY_ONE.equals(signUpsMapaStr)) {
                                                keyOne = l;
                                            }
                                            if (KEY_TWO.equals(signUpsMapaStr)) {
                                                keyTwo = l;
                                            }
                                        }
                                        //查找主键关联信息 从表内容查起
                                        for (int l = 1; l < signUps.size(); l++) {
                                            Map<Integer, Object> signUpsMapBody = signUps.get(l);
                                            String keyOneStr = (String) signUpsMapBody.get(keyOne);
                                            String keyTwoStr = (String) signUpsMapBody.get(keyTwo);

                                        }
                                        //组装输出list
                                        writeList.add(writeMap);
                                    }
                                }
                            }
                        } else {
                            continue;
                        }
                    }
                }
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
            e.printStackTrace();
        }

    }

    //招标材料
    public void tenderingMaterial() {

    }
}
