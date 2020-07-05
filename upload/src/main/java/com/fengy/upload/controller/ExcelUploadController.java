package com.fengy.upload.controller;


import com.fengy.upload.service.ExcelService;
import com.fengy.upload.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.net.URLEncoder;
import java.util.*;

/**
 * @Description: excel上传
 * @Author laoxu
 * @Date 2020/1/24 16:25
 **/
@Controller
@RequestMapping("/api/excel/")
public class ExcelUploadController {

    @Autowired
    private ExcelService service;

    private Map<String, Object> RESULT_DATA = new LinkedHashMap<>();

    private static String FILENAME = null;

    @RequestMapping("/")
    public String hello(){
        return "forward:index.html";
    }

    //接受文件上传
    @RequestMapping("/upload")
    @ResponseBody
    public Map<String,Object> uploadFile(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {
        Map<String,Object> map = new HashMap<>(16);

        FILENAME = file.getOriginalFilename();

        Map<String, Object> resMap = service.getExcelData(file);

        if (resMap == null){
            throw new Exception("文件为空！！");
        }

        RESULT_DATA.putAll(resMap);

        Map<String,Object> result = new HashMap<>();
        result.put("code",0);
        result.put("msg","");

        return result;
    }

    /**
     *  导出
     * @param response
     * @throws Exception
     */
    @GetMapping("/export")
    public void export(HttpServletResponse response)  throws Exception{
        if (RESULT_DATA == null){
            throw new Exception("请先上传文件！");
        }
        String filename = "--汇总金额表--.xls";

        response.setContentType("application/octet-stream");
        response.setHeader("Content-disposition",
                "attachment;filename=" + java.net.URLEncoder.encode(filename, "UTF-8"));

        Map<String, Object> resMap = RESULT_DATA;

        service.exportExcel(FILENAME,resMap,response);

        response.flushBuffer();

    }
}