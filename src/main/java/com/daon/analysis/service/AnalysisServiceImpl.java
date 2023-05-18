package com.daon.analysis.service;

import com.daon.analysis.dto.DuplicationData;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.*;

@Service
public class AnalysisServiceImpl implements AnalysisService {

    @Override
    public ArrayList<Map<String, String>> readData(MultipartFile file) throws Exception {
        ArrayList<Map<String, String>> dataList = new ArrayList<>();

        String extension = FilenameUtils.getExtension(file.getOriginalFilename());
        if (!extension.equals("xlsx") && !extension.equals("xls")) {
            throw new IOException("엑셀파일만 업로드 해주세요.");
        }

        Workbook workbook = null;

        if (extension.equals("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (extension.equals("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }

//        Sheet worksheet = workbook.getSheetAt(1);

        //시트명 불러오기
        for (int i = 0; i < workbook.getNumberOfSheets(); i++ ) {
            Map<String, String> sheetMap = new HashMap<>();
            sheetMap.put(String.valueOf(i), workbook.getSheetName(i));
            dataList.add(sheetMap);
        }
        return dataList;

    }

    @Override
    public File duplication(MultipartFile file, Integer sheetIndex) throws Exception {
        List<DuplicationData> duplicationDataList = new ArrayList<>();
        String extension = FilenameUtils.getExtension(file.getOriginalFilename());
        if (!extension.equals("xlsx") && !extension.equals("xls")) {
            throw new IOException("엑셀파일만 업로드 해주세요.");
        }
        Workbook workbook = null;
        if (extension.equals("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (extension.equals("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }

        Sheet worksheet = workbook.getSheetAt(sheetIndex);
        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) { // 4

            Row row = worksheet.getRow(i);
            DuplicationData duplicationData = new DuplicationData();
            try {
                if (null != row.getCell(5)) {
                    Cell mountaiNameCell = row.getCell(1);
                    Cell treeNameCell = row.getCell(2);

                    duplicationData.setPointName(mountaiNameCell.getStringCellValue());
                    duplicationData.setTreeName(treeNameCell.getStringCellValue());
                }
                if ((null == duplicationData.getTreeName() || "".equals(duplicationData.getTreeName())) && (null == duplicationData.getPointName() || "".equals(duplicationData.getPointName()))) {
                    continue;
                } else {
                    duplicationDataList.add(duplicationData);
                }
            } catch(Exception e) {
                e.printStackTrace();
            }

        }
        ArrayList<DuplicationData> resultList = new ArrayList<DuplicationData>(new HashSet<DuplicationData>(duplicationDataList));
        for (DuplicationData data: resultList) {
            System.out.println(data.toString());
        }
        System.out.println(resultList.size());
        return null;
    }
}
