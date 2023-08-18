package com.daon.analysis.service;

import com.daon.analysis.dto.DuplicationData;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
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
import java.util.stream.Collector;
import java.util.stream.Collectors;

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
        Set<String> pointSet = new HashSet<>();

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
        int totalCnt = 0;
        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) { // 4

            Row row = worksheet.getRow(i);
            DuplicationData duplicationData = new DuplicationData();

            //최초 cnt 1 설정
            duplicationData.setCnt(1);
            try {
                totalCnt ++;
                Cell mountaiNameCell = row.getCell(1);
                Cell treeNameCell = row.getCell(2);
                Cell diameterCell = row.getCell(3);

                String mountaiName = null == mountaiNameCell ?  "" : mountaiNameCell.getStringCellValue();
                if (StringUtils.isNotEmpty(mountaiName)) {
                    duplicationData.setPointName(mountaiName);
                    pointSet.add(mountaiName);

                    String treeName = null == mountaiNameCell ?  "" : treeNameCell.getStringCellValue();
                    duplicationData.setTreeName(treeName);
                    //직경 확인후 값 저장
                    if (null != diameterCell) {

                        Double diameter = diameterCell.getNumericCellValue();
                        if (null != diameter && 0 != diameter) {
                            duplicationData.setDiameter(diameter);
                        } else {
                            duplicationData.setDiameter(0D);
                        }

                        //포함여부 확인후 duplicationData
                        int idx = containsDuplicationData(duplicationDataList, duplicationData);
                        if (idx > 0) {
                            duplicationData.setDiameter(duplicationData.getDiameter() + duplicationDataList.get(idx).getDiameter());
                            duplicationData.setCnt(duplicationDataList.get(idx).getCnt() + 1);
                            duplicationDataList.remove(idx);
                        }

                    } else {
                        duplicationData.setDiameter(0D);
                    }

                    duplicationDataList.add(duplicationData);
                }

            } catch(Exception e) {
                e.printStackTrace();
            }

        }
        List<DuplicationData> resultList = new ArrayList<DuplicationData>(new HashSet<DuplicationData>(duplicationDataList));
        resultList = resultList.stream().sorted(Comparator.comparing(DuplicationData::getTreeName)).collect(Collectors.toList());
       for (DuplicationData data: resultList) {
            System.out.println(data.toString());
        }
        System.out.println(resultList.size());
        System.out.println(pointSet.size());
        for (String data: pointSet) {
            System.out.println(data);
        }
        //https://jforj.tistory.com/308 시트만들기
        return null;
    }

    /**  containsDuplicationData 중복여부 확인 **/
    public int containsDuplicationData(List<DuplicationData> list, DuplicationData duplicationData) {
        int idx = -1;
        for (int i = 0 ; i < list.size() ; i ++) {
            if (list.get(i).equals(duplicationData)) {
                idx = i;
                break;
            }
        }
        return idx;
    }
}
