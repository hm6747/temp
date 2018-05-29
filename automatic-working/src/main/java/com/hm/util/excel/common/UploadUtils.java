package com.hm.util.excel.common;

import com.hm.util.excel.base.BizResult;
import org.apache.commons.io.FileUtils;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.io.IOException;

/**
 * Created by Administrator on 2018/5/11 0011.
 */
public class UploadUtils {
    public BizResult upload(String fileName, HttpServletRequest request){
        MultipartHttpServletRequest mRequest = (MultipartHttpServletRequest) request;
        MultipartFile file = mRequest.getFile("file");
        if(file == null){
            return BizResult.createFailResult("请上传文件");
        }
        try {
            FileUtils.copyInputStreamToFile(file.getInputStream(),new File(this.getClass().getResource("/").getPath()+"/temp/"+fileName));
            return BizResult.createSuccessResult("上传成功",this.getClass().getResource("/").getPath()+"/temp/"+fileName);
        } catch (IOException e) {
            e.printStackTrace();
            return BizResult.createFailResult("上传失败");
        }
    }
    public String getPath(String fileName){
        return this.getClass().getResource("/").getPath()+"/temp/"+fileName;
    }
}
