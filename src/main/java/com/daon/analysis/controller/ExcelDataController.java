package com.daon.analysis.controller;

import com.daon.analysis.service.AnalysisService;
import jakarta.servlet.ServletOutputStream;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;

import java.awt.Color;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Map;

@Controller
@RequestMapping(value = "/excel")
public class ExcelDataController {

    @Autowired
    AnalysisService analysisService;

    @RequestMapping("/readExcelHeader.dor")
    public String read(Model model) throws Exception {
//        ArrayList<Map<String, String>> readDateList = analysisService.readData(file);
//        model.addAttribute("file", file);
//        model.addAttribute("sheerMap", readDateList);
        return "readExcelHeader";
    }

    @RequestMapping("/duplication.dor")
    public void duplication(@RequestParam Map<String,Object> paramMap, MultipartHttpServletRequest request, HttpServletResponse response) throws Exception {
        String sheetIndex = paramMap.get("sheetIndex").toString();
        String type = paramMap.get("type").toString();
        try {
            // 파일 읽어들이기
            MultipartFile file = null;
            Iterator<String> mIterator = request.getFileNames();
            if (mIterator.hasNext()) {
                file = request.getFile(mIterator.next());
            }
            String fileName = "";
            SXSSFWorkbook workbook = null;
            if ("1".equals(type) || "2".equals(type)) {
                if ("1".equals(type)) {
                    fileName = "duplicates";
                } else {
                    fileName = "treeDiameter";
                }
                workbook = analysisService.duplication(file, Integer.valueOf(sheetIndex), Integer.valueOf(type));
            } else if ("3".equals(type)) {
                fileName = "importantValue";
                workbook = analysisService.importantValue(file, Integer.valueOf(sheetIndex), Integer.valueOf(type));
            } else if ("4".equals(type)) {
                fileName = "diameter";
                workbook = analysisService.getDiameter(file, Integer.valueOf(sheetIndex), Integer.valueOf(type), "");
            }

            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
            ServletOutputStream servletOutputStream = response.getOutputStream();
            workbook.write(servletOutputStream);
            workbook.close();
            servletOutputStream.flush();
            servletOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @RequestMapping("/getDiameterForTreeName.dor")
    public void getDiameterForTreeName(@RequestParam Map<String,Object> paramMap, MultipartHttpServletRequest request, HttpServletResponse response) throws Exception {
        try {
            String sheetIndex = paramMap.get("sheetIndex").toString();
            String type = paramMap.get("type").toString();
            String treeName = paramMap.get("treeName").toString();
            // 파일 읽어들이기
            MultipartFile file = null;
            Iterator<String> mIterator = request.getFileNames();
            if (mIterator.hasNext()) {
                file = request.getFile(mIterator.next());
            }
            String fileName = "";
            SXSSFWorkbook workbook = null;

            fileName = "diameterForName";
            workbook = analysisService.getDiameter(file, Integer.valueOf(sheetIndex), Integer.valueOf(type), treeName);

            response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
            response.setHeader("Content-Disposition", "attachment;filename=" + fileName + ".xlsx");
            ServletOutputStream servletOutputStream = response.getOutputStream();
            workbook.write(servletOutputStream);
            workbook.close();
            servletOutputStream.flush();
            servletOutputStream.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}