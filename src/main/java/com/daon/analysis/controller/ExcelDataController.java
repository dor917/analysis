package com.daon.analysis.controller;

import com.daon.analysis.service.AnalysisService;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

@Controller
@RequestMapping(value = "/excel")
public class ExcelDataController {

    @Autowired
    AnalysisService analysisService;

    @RequestMapping("/readExcelHeader.dor")
    public String read(@RequestParam("file") MultipartFile file, Model model) throws Exception {
        ArrayList<Map<String, String>> readDateList = analysisService.readData(file);
        model.addAttribute("file", file);
        model.addAttribute("sheerMap", readDateList);
        return "readExcelHeader";
    }

    @RequestMapping("/duplication.dor")
        public void duplication(@RequestParam("sheetIndex") String sheetIndex, MultipartHttpServletRequest request) throws Exception {
        try {
            // 파일 읽어들이기
            MultipartFile file = null;
            Iterator<String> mIterator = request.getFileNames();
            if (mIterator.hasNext()) {
                file = request.getFile(mIterator.next());
            }
            analysisService.duplication(file, Integer.valueOf(sheetIndex));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }



}