package com.daon.analysis.service;

import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.Map;

public interface AnalysisService {
    public ArrayList<Map<String, String>> readData(MultipartFile mf) throws Exception;
}
