package com.daon.analysis.service;

import com.daon.analysis.dto.DuplicationData;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;
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
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Map<String, String> sheetMap = new HashMap<>();
            sheetMap.put(String.valueOf(i), workbook.getSheetName(i));
            dataList.add(sheetMap);
        }
        return dataList;

    }

    @Override
    public SXSSFWorkbook duplication(MultipartFile file, Integer sheetIndex) throws Exception {
        final String fileName = "temp";
        List<DuplicationData> duplicationDataList = new ArrayList<>();
        Map<String, DuplicationData> duplicationDataMap = new HashMap<>();
        double allCnt = 0; //평균 구하기용 전체 카운트
        String extension = FilenameUtils.getExtension(file.getOriginalFilename());
        Set<String> pointSet = new HashSet<>();
        Set<String> treeSet = new HashSet<>();
        Cell diameterCell = null;

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
                Cell mountaiNameCell = row.getCell(1);
                Cell treeNameCell = row.getCell(2);
                diameterCell = row.getCell(3);

                String mountaiName = null == mountaiNameCell ? "" : mountaiNameCell.getStringCellValue();
                if (StringUtils.isNotEmpty(mountaiName)) {
                    duplicationData.setPointName(mountaiName);
                    pointSet.add(mountaiName);

                    String treeName = null == mountaiNameCell ? "" : treeNameCell.getStringCellValue();
                    duplicationData.setTreeName(treeName);
                    treeSet.add(treeName);
                    //직경 확인후 값 저장
                    if (null != diameterCell) {
                        Double diameter = null;
                        // 면적 String 형식으로 저장되었을 경우 형변환
                        if (CellType.STRING ==diameterCell.getCellType()) {
                            if (null != diameterCell.getStringCellValue()) {
                                diameter = Double.valueOf(diameterCell.getStringCellValue());
                            } else {
                                diameter = 0D;
                            }

                        } else {
                            diameter = diameterCell.getNumericCellValue();
                        }

                        if (null != diameter && 0D < diameter) {
                            duplicationData.setDiameter(diameter);
                        } else {
                            duplicationData.setDiameter(0D);
                        }

                        //포함여부 확인후 duplicationData

                        int idx = containsDuplicationData(duplicationDataList, duplicationData);
                        if (idx >= 0) {
                            if (0D < duplicationData.getDiameter()) {
                                duplicationData.setDiameter(duplicationData.getDiameter() + duplicationDataList.get(idx).getDiameter());
                                duplicationData.setCnt(duplicationDataList.get(idx).getCnt() + 1);
                                duplicationDataList.remove(idx);
                            }
                        } else {
                            if (0D < duplicationData.getDiameter()) {
                                duplicationData.setCnt(1);
                            }
                        }
//난티나무
                    } else {
                        duplicationData.setDiameter(0D);
                    }
                    if (duplicationData.getCnt() > 0 && duplicationData.getDiameter() > 0D ) {
                        allCnt ++;
                        duplicationDataList.add(duplicationData);
                        duplicationDataMap.put(duplicationData.getPointName() + "^%^" + duplicationData.getTreeName(), duplicationData);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
        List<DuplicationData> resultList = new ArrayList<DuplicationData>(new HashSet<DuplicationData>(duplicationDataList));
        resultList = resultList.stream().sorted(Comparator.comparing(DuplicationData::getTreeName)).collect(Collectors.toList());
        workbook.close();
        //https://jforj.tistory.com/308 시트만들기

        /**
         * excel sheet 생성
         */
        SXSSFWorkbook writeWorkbook = new SXSSFWorkbook();
        Sheet sheet = writeWorkbook.createSheet("Sheet1"); // 엑셀 sheet 이름
        sheet.setDefaultColumnWidth(28); // 디폴트 너비 설정

        /**
         * header font style
         */
        XSSFFont headerXSSFFont = (XSSFFont) writeWorkbook.createFont();
        headerXSSFFont.setColor(new XSSFColor(new java.awt.Color(0)));

        /**
         * header cell style
         */
        XSSFCellStyle headerXssfCellStyle = (XSSFCellStyle) writeWorkbook.createCellStyle();

        // 테두리 설정
        headerXssfCellStyle.setBorderLeft(BorderStyle.THIN);
        headerXssfCellStyle.setBorderRight(BorderStyle.THIN);
        headerXssfCellStyle.setBorderTop(BorderStyle.THIN);
        headerXssfCellStyle.setBorderBottom(BorderStyle.THIN);

        // 배경 설정
//        headerXssfCellStyle.setFillForegroundColor(new XSSFColor(new byte[]{(byte) 34, (byte) 37, (byte) 41}));
//        headerXssfCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerXssfCellStyle.setFont(headerXSSFFont);

        /**
         * body cell style
         */
        XSSFCellStyle bodyXssfCellStyle = (XSSFCellStyle) writeWorkbook.createCellStyle();

        // 테두리 설정
        bodyXssfCellStyle.setBorderLeft(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderRight(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderTop(BorderStyle.THIN);
        bodyXssfCellStyle.setBorderBottom(BorderStyle.THIN);

        /**
         * header data
         */
        int rowCount = 0; // 데이터가 저장될 행
        Row headerRow = null;
        Cell headerCell = null;
        headerRow = sheet.createRow(rowCount++);

        List<String> headerList = new ArrayList<>(pointSet);
        headerCell = headerRow.createCell(0);
        headerCell.setCellValue("수종명");
        headerCell.setCellStyle(headerXssfCellStyle);
        for (int i = 0; i < headerList.size(); i++) {
            headerCell = headerRow.createCell(i + 1);
            headerCell.setCellValue(headerList.get(i)); // 데이터 추가
            headerCell.setCellStyle(headerXssfCellStyle); // 스타일 추가
        }
        headerCell = headerRow.createCell(headerList.size() + 1);
        headerCell.setCellValue("평균");
        headerCell.setCellStyle(headerXssfCellStyle);
        headerCell = headerRow.createCell(headerList.size() + 2);
        headerCell.setCellValue("%");
        headerCell.setCellStyle(headerXssfCellStyle);

        /**
         * body data
         */
        Row bodyRow = null;
        Cell bodyCell = null;
        Cell avgCell = null;
        List<String> bodyList = new ArrayList<>(treeSet);

        // 수종별 표본고점 데이터 세팅
        for (String treeName : bodyList) {
            double average = 0;
            bodyRow = sheet.createRow(rowCount++);
            bodyCell = bodyRow.createCell(0);
            bodyCell.setCellValue(treeName); // 데이터 추가
            for (int i = 0; i < headerList.size(); i++) {
                DuplicationData data = duplicationDataMap.get(headerList.get(i) + "^%^" + treeName);
                Cell bodyCntCell = bodyRow.createCell(i + 1);

                if (null != data) {
                    bodyCntCell.setCellValue(data.getCnt());
                    average += data.getCnt();
                } else {
                    bodyCntCell.setCellValue(0);
                }


                bodyCntCell.setCellStyle(bodyXssfCellStyle);
            }
            // 평균 셀
            avgCell = bodyRow.createCell(headerList.size() + 1);
            avgCell.setCellValue(average / headerList.size());
            avgCell.setCellStyle(bodyXssfCellStyle);
            avgCell = bodyRow.createCell(headerList.size() + 2);
            avgCell.setCellValue((average / allCnt) * 100);
            avgCell.setCellStyle(bodyXssfCellStyle);
            bodyCell.setCellStyle(bodyXssfCellStyle); // 스타일 추가
        }

        return writeWorkbook;
    }

    /**
     * containsDuplicationData 중복여부 확인
     **/
    public int containsDuplicationData(List<DuplicationData> list, DuplicationData duplicationData) {
        int idx = -1;
        for (int i = 0; i < list.size(); i++) {
            if (list.get(i).equals(duplicationData)) {
                idx = i;
                break;
            }
        }
        return idx;
    }
}
