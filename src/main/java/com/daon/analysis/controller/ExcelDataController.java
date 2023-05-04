package com.daon.analysis.controller;

import com.daon.analysis.service.AnalysisService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.Map;

@Controller
@RequestMapping(value = "excel")
public class ExcelDataController {

    @Autowired
    AnalysisService analysisService;

    @RequestMapping("/read")
    public String read(@RequestParam("file") MultipartFile file, Model model) throws Exception {
        ArrayList<Map<String, String>> readDateList = analysisService.readData(file);


        return "ㅁㅁㅁㅁ";
    }



}