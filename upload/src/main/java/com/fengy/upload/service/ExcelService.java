package com.fengy.upload.service;

import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.util.Map;

public interface ExcelService {

    public Map<String,Object> getExcelData(MultipartFile file);

    public String exportExcel(String olename, Map<String, Object> resMap, HttpServletResponse response);
}
