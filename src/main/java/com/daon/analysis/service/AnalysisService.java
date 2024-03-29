package com.daon.analysis.service;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.Map;

public interface AnalysisService {
    public ArrayList<Map<String, String>> readData(MultipartFile mf) throws Exception;

    public SXSSFWorkbook duplication (MultipartFile mf, Integer sheetIndex, Integer type) throws Exception;

    public SXSSFWorkbook importantValue(MultipartFile file, Integer sheetIndex, Integer integer) throws Exception;

    public SXSSFWorkbook getDiameter(MultipartFile file, Integer integer, Integer integer1, String treeName) throws Exception;
}
