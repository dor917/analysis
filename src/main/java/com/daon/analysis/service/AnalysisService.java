package com.daon.analysis.service;

import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.util.ArrayList;
import java.util.Map;

public interface AnalysisService {
    public ArrayList<Map<String, String>> readData(MultipartFile mf) throws Exception;
    public File duplication (MultipartFile mf, Integer sheetIndex) throws Exception;
}
