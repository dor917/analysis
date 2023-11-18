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
import java.math.BigDecimal;
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


    /**
     * 수목별 중복수 , 흉고 단면적 값 구하기
     */
    @Override
    public SXSSFWorkbook duplication(MultipartFile file, Integer sheetIndex, Integer type) throws Exception {

        Workbook workbook = getWorkbook(file);

        SXSSFWorkbook writeWorkbook = new SXSSFWorkbook();

        if (1 == type || 2 == type) { //수목별 중복수 구하기
            Row getheaderRow = workbook.getSheetAt(sheetIndex).getRow(0);
            Map<String, Integer> diameterMap = findDiameterColumns(getheaderRow);  //흉고직경 로우 가져오기
            Map<String, Integer> sortedMap = new TreeMap<>(diameterMap);
            for ( String diameterKey : sortedMap.keySet() ) {
                List<DuplicationData> duplicationDataList = new ArrayList<>();
                Map<String, DuplicationData> duplicationDataMap = new HashMap<>();
                Map<String, Double> allCntMap = new HashMap<>();


                Set<String> moutainSet = new HashSet<>();
                Map<String, Set<String>> pointSetMap = new HashMap<>();
                Map<String, Set<String>> treeSetMap = new HashMap<>();

                getExcelData(workbook, sheetIndex, diameterMap.get(diameterKey), type, duplicationDataList, allCntMap, duplicationDataMap, moutainSet, pointSetMap, treeSetMap, diameterKey);
                for (String moutainKey : moutainSet) {
                    /**
                     * excel sheet 생성
                     */

                    Sheet sheet = writeWorkbook.createSheet(moutainKey+ "(" + diameterKey + ")"); // 엑셀 sheet 이름
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
                    headerXssfCellStyle.setAlignment(HorizontalAlignment.CENTER);

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

                    List<String> headerList = new ArrayList<>(pointSetMap.get(moutainKey));
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
                    List<String> bodyList = new ArrayList<>(treeSetMap.get(moutainKey));

                    if (1 == type) {   // 수종별 표본고점 데이터 세팅

                        for (String treeName : bodyList) {
                            double average = 0;
                            bodyRow = sheet.createRow(rowCount++);
                            bodyCell = bodyRow.createCell(0);
                            bodyCell.setCellValue(treeName); // 데이터 추가
                            for (int i = 0; i < headerList.size(); i++) {
                                DuplicationData data = duplicationDataMap.get(moutainKey + "^%^" + diameterKey + "^%^" + headerList.get(i) + "^%^" + treeName);
                                Cell bodyCntCell = bodyRow.createCell(i + 1);

                                if (null != data) {
                                    bodyCntCell.setCellValue(data.getCnt());
                                    average += data.getCnt();
                                } else {
                                    bodyCntCell.setCellValue(0);
                                }


                                bodyCntCell.setCellStyle(bodyXssfCellStyle);
                            }
                            double percent =  allCntMap.get(moutainKey + "^%^" + diameterKey) == null ? 0D : average / allCntMap.get(moutainKey + "^%^" +  diameterKey) * 100;
                            // 평균 셀
                            avgCell = bodyRow.createCell(headerList.size() + 1);
                            avgCell.setCellValue(average / headerList.size());
                            avgCell.setCellStyle(bodyXssfCellStyle);
                            avgCell = bodyRow.createCell(headerList.size() + 2);
                            avgCell.setCellValue(percent);
                            avgCell.setCellStyle(bodyXssfCellStyle);
                            bodyCell.setCellStyle(bodyXssfCellStyle); // 스타일 추가
                        }
                    } else {// 수종별 표본고점 데이터 세팅
                        /**
                         * body data
                         */
                        for (String treeName : bodyList) {
                            double average = 0;
                            bodyRow = sheet.createRow(rowCount++);
                            bodyCell = bodyRow.createCell(0);
                            bodyCell.setCellValue(treeName); // 데이터 추가
                            for (int i = 0; i < headerList.size(); i++) {
                                DuplicationData data = duplicationDataMap.get(moutainKey + "^%^" + diameterKey + "^%^" + headerList.get(i) + "^%^" + treeName);
                                Cell bodyCntCell = bodyRow.createCell(i + 1);

                                if (null != data) {
                                    bodyCntCell.setCellValue(data.getDiameter() * 0.0001);
                                    average += data.getDiameter() * 0.0001;
                                } else {
                                    bodyCntCell.setCellValue(0);
                                }


                                bodyCntCell.setCellStyle(bodyXssfCellStyle);
                            }
                            double percent =  allCntMap.get(moutainKey + "^%^" + diameterKey) == null ? 0D : average / (allCntMap.get(moutainKey + "^%^" +  diameterKey) * 0.0001) * 100;
                            // 평균 셀
                            avgCell = bodyRow.createCell(headerList.size() + 1);
                            avgCell.setCellValue(new BigDecimal(average / headerList.size()).setScale(4, BigDecimal.ROUND_FLOOR).doubleValue());
                            avgCell.setCellStyle(bodyXssfCellStyle);
                            avgCell = bodyRow.createCell(headerList.size() + 2);
                            avgCell.setCellValue(new BigDecimal(percent).setScale(4, BigDecimal.ROUND_FLOOR).doubleValue());
                            avgCell.setCellStyle(bodyXssfCellStyle);
                            bodyCell.setCellStyle(bodyXssfCellStyle); // 스타일 추가
                        }
                    }
                }

            }

        }
        workbook.close();

        return writeWorkbook;
    }

    /**
     * 중요치 구하기
     */
    @Override
    public SXSSFWorkbook importantValue(MultipartFile file, Integer sheetIndex, Integer integer) throws Exception{
        Workbook workbook = getWorkbook(file);
        SXSSFWorkbook writeWorkbook = new SXSSFWorkbook();
        Sheet worksheet = workbook.getSheetAt(sheetIndex);
        int treeIndex = 0; //수종명 index
        int pointIndex = 0; //표본점번호 index
        int moutainIndex = 0; //산지명 index

        Map<String, Double> treeYearAllCountMap = new HashMap<>();
        Map<String, Double> allDiameterMap = new HashMap<>();
        Set<String> checkSet = new HashSet<>();
        Map<String, Integer> emergenceMap = new HashMap<>();
        Map<String, Double> allEmergenceMap = new HashMap<>();
        Row getHeaderRow = workbook.getSheetAt(sheetIndex).getRow(0);
        Map<String, Integer> diameterMap = findDiameterColumns(getHeaderRow); // 년도
        Map<String, Integer> treeYearCountMap = new HashMap<>();
        Map<String, Double> treeYeardiameterMap = new HashMap<>();
        Set<String> moutainSet = new HashSet<>();
        Set<String> treeSet = new HashSet<>();

        for (int i = 0; i < getHeaderRow.getPhysicalNumberOfCells(); i++) {
            String rowText = getHeaderRow.getCell(i).getStringCellValue().trim();
            if (rowText.contains("수종명")) {
                treeIndex = i;
            } else if (rowText.contains("표본점번호")) {
                pointIndex = i;
            } else if (rowText.contains("산지명")) {
                moutainIndex = i;
            }
        }
        //map 수종 년도별 수종수
        //map2 주송 년도별 흉고 단면적
        for ( String diameterKey : diameterMap.keySet() ) {
            for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
                Row row = worksheet.getRow(i);
                Cell moutainNameCell = row.getCell(moutainIndex);
                Cell pointNameCell = row.getCell(pointIndex);
                Cell treeNameCell = row.getCell(treeIndex);
                Cell diameterCell = row.getCell(diameterMap.get(diameterKey));
                String pointName = null == pointNameCell ? "" : pointNameCell.getStringCellValue();
                if (StringUtils.isNotEmpty(pointName)) {

                    //산이름
                    String moutainName = null == moutainNameCell ? "" : moutainNameCell.getStringCellValue();
                    String treeName = null == pointNameCell ? "" : treeNameCell.getStringCellValue();
                    if (StringUtils.isNotEmpty(treeName)) {
                        treeSet.add(treeName);
                    }
                    if (StringUtils.isNotEmpty(moutainName)) {
                        moutainSet.add(moutainName);
                    }
                    String mapKey = diameterKey + "^%^" + moutainName + "^%^" + treeName;
                    String allCntMapKey = diameterKey + "^%^" + moutainName;
                    String checkKey = diameterKey + "^%^" + moutainName + "^%^" + treeName + "^%^" + pointName;


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
                            if (!checkSet.contains(checkKey)) {
                                //년도, 산, 나무별 카운트
                                int cnt = emergenceMap.get(mapKey) == null ? 1 : emergenceMap.get(mapKey) + 1;
                                emergenceMap.put(mapKey, cnt);
                                // 전체 카운트
                                double allCnt = allEmergenceMap.get(allCntMapKey)  == null ? 1D : allEmergenceMap.get(allCntMapKey) + 1D;
                                allEmergenceMap.put(allCntMapKey, allCnt);
                                checkSet.add(checkKey);
                            }

                            //년도, 산, 나무별 카운트
                            int cnt = treeYearCountMap.get(mapKey) == null ? 1 : treeYearCountMap.get(mapKey) + 1;
                            treeYearCountMap.put(mapKey, cnt);
                            // 전체 카운트
                            double allCnt = treeYearAllCountMap.get(allCntMapKey)  == null ? 1 : treeYearAllCountMap.get(allCntMapKey) + 1;
                            treeYearAllCountMap.put(allCntMapKey, allCnt);
                            diameter = ((diameter / 2) * (diameter / 2)) * 3.14 * 0.0001;  // 흉고 단면적 값 구하기 일 경우 흉고직경 값에 파이알 제곱
                            //년도, 산, 나무별 흉고직경 더하기
                            double diameterSum = treeYeardiameterMap.get(mapKey) == null ? diameter : treeYeardiameterMap.get(mapKey) + diameter;
                            treeYeardiameterMap.put(mapKey, diameterSum);

                            // 전체 흉고직경
                            double allDiameter = allDiameterMap.get(allCntMapKey)  == null ? diameter : allDiameterMap.get(allCntMapKey) + diameter;
                            allDiameterMap.put(allCntMapKey, allDiameter);
                        }

                    }

                }
            }
        }
        Map<String, Integer> sortedMap = new TreeMap<>(diameterMap);
        for ( String diameterKey : sortedMap.keySet() ) {
            for (String moutainKey : moutainSet) {

                Sheet sheet = writeWorkbook.createSheet(moutainKey+ "(" + diameterKey +")" ); // 엑셀 sheet 이름
                sheet.setDefaultColumnWidth(35); // 디폴트 너비 설정
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
                headerXssfCellStyle.setAlignment(HorizontalAlignment.CENTER); // 가운데 정렬 (가로 기준)

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
                int headerCellCount = 0; // 데이터가 저장될 행
                Row headerRow = null;
                Cell headerCell = null;
                headerRow = sheet.createRow(rowCount ++);
                headerCell = headerRow.createCell(headerCellCount ++);
                headerCell.setCellValue("수종명");
                headerCell.setCellStyle(headerXssfCellStyle);


                headerCell = headerRow.createCell(headerCellCount ++);
                headerCell.setCellValue("출연지점");
                headerCell.setCellStyle(headerXssfCellStyle);
                headerCell = headerRow.createCell(headerCellCount ++);
                headerCell.setCellValue("%");
                headerCell.setCellStyle(headerXssfCellStyle);
                headerCellCount ++;
                headerCell = headerRow.createCell(headerCellCount ++);
                headerCell.setCellValue("수종명");
                headerCell.setCellStyle(headerXssfCellStyle);
                headerCell = headerRow.createCell(headerCellCount);
                headerCell.setCellValue("중요치");
                headerCell.setCellStyle(headerXssfCellStyle);

                Row bodyRow = null;
                Cell bodyCell = null;
                Cell avgCell = null;
                Cell bodyCntCell = null;

                for (String treeName : treeSet) {
                    String dmtKey = diameterKey + "^%^" + moutainKey + "^%^" + treeName;
                    String dmKey = diameterKey + "^%^" + moutainKey;
                    double treeYearCount = treeYearCountMap.get(dmtKey) == null ? 0 : treeYearCountMap.get(dmtKey);
                    double treeYearAllCount = treeYearAllCountMap.get(dmKey) == null ? 0 : treeYearAllCountMap.get(dmKey);
                    double treeYeardiameter = treeYeardiameterMap.get(dmtKey) == null ? 0 : treeYeardiameterMap.get(dmtKey);
                    double allDiameter = allDiameterMap.get(dmKey) == null ? 0 : allDiameterMap.get(dmKey);

                    int emergenceCnt = emergenceMap.get(dmtKey) == null ? 0 : emergenceMap.get(dmtKey); // 년도별 산별 중요치
                    double emergenceAllCnt = allEmergenceMap.get(dmKey) == null ? 0D : allEmergenceMap.get(dmKey); // 년도별 산별 중요치
                    if (emergenceCnt > 0) {
                        double cntAvgValue = (emergenceCnt / emergenceAllCnt) * 100;
                        double pointCntAvgValue =(treeYearCount / treeYearAllCount) * 100;
                        double diameterAvgValue = (treeYeardiameter / allDiameter) * 100;
                        double importantValue = (pointCntAvgValue + diameterAvgValue + cntAvgValue) / 3;
                        bodyRow = sheet.createRow(rowCount++);
                        bodyCell = bodyRow.createCell(0);
                        bodyCell.setCellValue(treeName); // 데이터 추가
                        bodyCell.setCellStyle(bodyXssfCellStyle);
                        bodyCntCell = bodyRow.createCell( 1);
                        bodyCntCell.setCellValue(emergenceCnt);
                        bodyCntCell.setCellStyle(bodyXssfCellStyle);
                        avgCell = bodyRow.createCell( 2);
                        avgCell.setCellValue(new BigDecimal(cntAvgValue).setScale(4, BigDecimal.ROUND_FLOOR).doubleValue());
                        avgCell.setCellStyle(bodyXssfCellStyle);
                        bodyCell = bodyRow.createCell(4);
                        bodyCell.setCellValue(treeName); // 데이터 추가
                        bodyCell.setCellStyle(bodyXssfCellStyle);


                        bodyCell = bodyRow.createCell(5);
                        bodyCell.setCellValue(new BigDecimal(importantValue).setScale(4, BigDecimal.ROUND_FLOOR).doubleValue()); // 데이터 추가
                        bodyCell.setCellStyle(bodyXssfCellStyle);

                    }


                }
            }



        }
        workbook.close();

        return writeWorkbook;
    }

    @Override
    public SXSSFWorkbook getDiameter(MultipartFile file, Integer sheetIndex, Integer integer, String treeName) throws Exception {
        Workbook workbook = getWorkbook(file);
        SXSSFWorkbook writeWorkbook = new SXSSFWorkbook();
        Sheet worksheet = workbook.getSheetAt(sheetIndex);
        int treeIndex = 0; //수종명 index
        int pointIndex = 0; //표본점번호 index
        int moutainIndex = 0; //산지명 index
        int stateIndex = 0; //수간상태 index
        int allCnt = 0 ;
        Row getHeaderRow = workbook.getSheetAt(sheetIndex).getRow(0);
        Map<String, Integer> yearMap = findDiameterColumns(getHeaderRow); // 년도


        try {
            for (int i = 0; i < getHeaderRow.getPhysicalNumberOfCells(); i++) {
                String rowText = getHeaderRow.getCell(i).getStringCellValue().trim();
                if (rowText.contains("수종명")) {
                    treeIndex = i;
                } else if (rowText.contains("표본점번호")) {
                    pointIndex = i;
                } else if (rowText.contains("산지명")) {
                    moutainIndex = i;
                }
            }
            Map<String, Integer> sortedMap = new TreeMap<>(yearMap);
            for ( String yearKey : sortedMap.keySet() ) {
                Set<String> moutainSet = new HashSet<>();
                Set<Double> diameterSet = new TreeSet<>();
                Set<String> pointSet = new TreeSet<>();
                Map<String, Integer> diameterPointCntMap= new HashMap<>();
                for (int i = 0; i < getHeaderRow.getPhysicalNumberOfCells(); i++) {
                    String rowText = getHeaderRow.getCell(i).getStringCellValue().replace(System.getProperty("line.separator").toString(), "");
                    if (rowText.contains("수간상태("+yearKey+")") || rowText.contains("수간 상태("+yearKey+")")) {
                        stateIndex = i;
                        break;
                    }
                }
                for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) {
                    Row row = worksheet.getRow(i);
                    Cell moutainNameCell = row.getCell(moutainIndex);
                    Cell pointNameCell = row.getCell(pointIndex);
                    Cell treeNameCell = row.getCell(treeIndex);
                    Cell diameterCell = row.getCell(yearMap.get(yearKey));
                    Cell stateCell = row.getCell(stateIndex);
                    String moutainName = null == moutainNameCell ? "" : moutainNameCell.getStringCellValue();
                    String cellTreeName = null == treeNameCell ? "" : treeNameCell.getStringCellValue();


                    if (StringUtils.isNotEmpty(moutainName)) {
                        moutainSet.add(moutainName);
                    }
                    if (StringUtils.isNotEmpty(treeName)) {
                        if (StringUtils.isNotEmpty(cellTreeName)) {
                            if (!cellTreeName.equals(treeName)) {
                                continue;
                            }
                        }
                    }

                    Double state = 0D; // 상태
                    if (null != stateCell) {
                        // 면적 String 형식으로 저장되었을 경우 형변환
                        if (CellType.STRING == stateCell.getCellType()) {
                            if (null != stateCell.getStringCellValue()) {
                                state = Double.valueOf(stateCell.getStringCellValue());
                            } else {
                                state = 0D;
                            }

                        } else {
                            state = stateCell.getNumericCellValue();
                        }
                    }

                    if (state < 6D) {
                        String pointName = null == pointNameCell ? "" : pointNameCell.getStringCellValue();
                        //직경 확인후 값 저장
                        Double diameter = 0D;
                        if (null != diameterCell) {
                            // 면적 String 형식으로 저장되었을 경우 형변환
                            if (CellType.STRING == diameterCell.getCellType()) {
                                if (null != diameterCell.getStringCellValue()) {
                                    diameter = Double.valueOf(diameterCell.getStringCellValue());
                                } else {
                                    diameter = 0D;
                                }

                            } else {
                                diameter = diameterCell.getNumericCellValue();
                            }
                        }
                       if (diameter > 0D) {
                           String key =  pointName + "^%^" + diameter;
                           if (StringUtils.isNotEmpty(pointName)) {
                               pointSet.add(pointName);
                           }
                           diameterSet.add(diameter);

                           int cnt = diameterPointCntMap.get(key) == null ? 1 : diameterPointCntMap.get(key) + 1;
                           diameterPointCntMap.put(key, cnt);
                           allCnt ++;
                       }
                    }

                }
                for (String moutainKey: moutainSet) {

                    Sheet sheet = writeWorkbook.createSheet(moutainKey+ "(" + yearKey +")" ); // 엑셀 sheet 이름
                    sheet.setDefaultColumnWidth(35); // 디폴트 너비 설정
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
                    headerXssfCellStyle.setAlignment(HorizontalAlignment.CENTER); // 가운데 정렬 (가로 기준)

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
                    int headerCellCount = 0; // 데이터가 저장될 행
                    Row headerRow = null;
                    Cell headerCell = null;
                    headerRow = sheet.createRow(rowCount ++);
                    headerCell = headerRow.createCell(headerCellCount ++);
                    headerCell.setCellValue("흉고직경");
                    headerCell.setCellStyle(headerXssfCellStyle);

                    for (String pointKey : pointSet) {
                        headerCell = headerRow.createCell(headerCellCount ++);
                        headerCell.setCellValue(pointKey);
                        headerCell.setCellStyle(headerXssfCellStyle);
                    }
                    headerCell = headerRow.createCell(headerCellCount ++);
                    headerCell.setCellValue("평균");
                    headerCell.setCellStyle(headerXssfCellStyle);

                    Row bodyRow = null;
                    Cell bodyCell = null;
                    Cell avgCell = null;
                    Cell bodyCntCell = null;
                    for (double diameterKey : diameterSet) {
                        int cellCnt = 0;
                        bodyRow = sheet.createRow(rowCount++);
                        bodyCell = bodyRow.createCell(cellCnt++);
                        bodyCell.setCellValue(diameterKey); // 데이터 추가
                        bodyCell.setCellStyle(bodyXssfCellStyle);
                        double sumCnt = 0;
                        for (String pointKey : pointSet) {
                            String key =  pointKey + "^%^" + diameterKey;
                            int cntValue = diameterPointCntMap.get( pointKey + "^%^" + diameterKey) == null ? 0 : diameterPointCntMap.get( pointKey + "^%^" + diameterKey) * 25;
                            sumCnt += cntValue;
                            bodyCell = bodyRow.createCell(cellCnt++);
                            bodyCell.setCellValue(cntValue); // 데이터 추가
                            bodyCell.setCellStyle(bodyXssfCellStyle);
                        }
                        bodyCell = bodyRow.createCell(cellCnt++);
                        bodyCell.setCellValue(new BigDecimal(sumCnt / pointSet.size()).setScale(4, BigDecimal.ROUND_FLOOR).doubleValue()); // 데이터 추가
                        bodyCell.setCellStyle(bodyXssfCellStyle);

                    }

                }

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        workbook.close();

        return writeWorkbook;
    }


    /**
     *  엑셀 파일 데이터 읽기 return
     **/

    private void getExcelData(Workbook workbook, int sheetIndex, int diameterRow, int type, List<DuplicationData> duplicationDataList, Map<String, Double> allCntMap, Map<String, DuplicationData> duplicationDataMap,
                             Set<String> moutainSet, Map<String, Set<String>> pointSetMap, Map<String, Set<String>> treeSetMap, String diameterKey) {
        Sheet worksheet = workbook.getSheetAt(sheetIndex);
        Row headerRow = workbook.getSheetAt(sheetIndex).getRow(0);
        int treeIndex = 0; //수종명 index
        int pointIndex = 0; //표본점번호 index
        int moutainIndex = 0; //산지명 index

        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            String rowText = headerRow.getCell(i).getStringCellValue().trim();
            if (rowText.contains("수종명")) {
                treeIndex = i;
            } else if (rowText.contains("표본점번호")) {
                pointIndex = i;
            } else if (rowText.contains("산지명")) {
                moutainIndex = i;
            }
        }
        Cell diameterCell = null;
        for (int i = 1; i < worksheet.getPhysicalNumberOfRows(); i++) { // 4

            Row row = worksheet.getRow(i);
            DuplicationData duplicationData = new DuplicationData();

            try {
                Cell moutainNameCell = row.getCell(moutainIndex);
                Cell pointNameCell = row.getCell(pointIndex);
                Cell treeNameCell = row.getCell(treeIndex);
                diameterCell = row.getCell(diameterRow);

                String pointName = null == pointNameCell ? "" : pointNameCell.getStringCellValue();
                if (StringUtils.isNotEmpty(pointName)) {
                    duplicationData.setYear(diameterKey);
                    //산이름
                    String moutainName = null == moutainNameCell ? "" : moutainNameCell.getStringCellValue();
                    duplicationData.setMoutinaName(moutainName);
                    moutainSet.add(moutainName);

                    duplicationData.setPointName(pointName);
                    if (null == pointSetMap.get(moutainName)) {
                        Set<String> pointSet = new HashSet<>();
                        pointSet.add(pointName);
                        pointSetMap.put(moutainName, pointSet);
                    } else {
                        Set<String> pointSet = pointSetMap.get(moutainName);
                        pointSet.add(pointName);
                        pointSetMap.put(moutainName, pointSet);
                    }


                    String treeName = null == pointNameCell ? "" : treeNameCell.getStringCellValue();
                    duplicationData.setTreeName(treeName);

                    if (null == treeSetMap.get(moutainName)) {
                        Set<String> treeSet = new HashSet<>();
                        treeSet.add(treeName);
                        treeSetMap.put(moutainName, treeSet);
                    } else {
                        Set<String> treeSet = treeSetMap.get(moutainName);
                        treeSet.add(treeName);
                        treeSetMap.put(moutainName, treeSet);
                    }

                    //직경 확인후 값 저장
                    Double diameter = null;
                    if (null != diameterCell) {
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
                            if (2 == type) { // 흉고 단면적 값 구하기 일 경우 흉고직경 값에 파이알 제곱
                                diameter = (diameter / 2) * (diameter / 2) * 3.14;
                            }
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

                    } else {
                        duplicationData.setDiameter(0D);
                    }
                    if (duplicationData.getCnt() > 0 && duplicationData.getDiameter() > 0D ) {
                        if (1 == type) {
                            if (allCntMap.get(duplicationData.getMoutinaName() + "^%^" + diameterKey) == null) {
                                allCntMap.put(duplicationData.getMoutinaName() + "^%^" + diameterKey, 1D);
                            } else {
                                allCntMap.put(duplicationData.getMoutinaName() + "^%^" + diameterKey, allCntMap.get(duplicationData.getMoutinaName() + "^%^"  + diameterKey) + 1);
                            }
                        } else if (2 == type) {
                            if (allCntMap.get(duplicationData.getMoutinaName() + "^%^" + diameterKey) == null) {
                                allCntMap.put(duplicationData.getMoutinaName() + "^%^" + diameterKey, diameter);
                            } else {
                                allCntMap.put(duplicationData.getMoutinaName() + "^%^" + diameterKey, allCntMap.get(duplicationData.getMoutinaName() + "^%^"  + diameterKey) + diameter);
                            }
                        }

                        duplicationDataList.add(duplicationData);
                        duplicationDataMap.put(duplicationData.getMoutinaName() + "^%^" + diameterKey + "^%^" + duplicationData.getPointName() + "^%^" + duplicationData.getTreeName(), duplicationData);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }

        }
    }
    /**
     * 흉고직경 컬럼 찾기
     */
    private Map<String, Integer> findDiameterColumns(Row headerRow) {
        Map<String, Integer> diameterMap = new HashMap<>();
        for (int i = 0; i < headerRow.getPhysicalNumberOfCells(); i++) {
            String rowText = headerRow.getCell(i).getStringCellValue().trim();
            if (rowText.contains("흉고직경")) {
                String intStr = rowText.replaceAll("[^0-9]", "");
                diameterMap.put(intStr, i);
            }
        }
        return diameterMap;
    }

    /**
     * containsDuplicationData 중복여부 확인
     **/
    private int containsDuplicationData(List<DuplicationData> list, DuplicationData duplicationData) {
        int idx = -1;
        for (int i = 0; i < list.size(); i++) {
            if (list.get(i).equals(duplicationData)) {
                idx = i;
                break;
            }
        }
        return idx;
    }

    private Workbook getWorkbook (MultipartFile file) throws Exception{
        String extension = FilenameUtils.getExtension(file.getOriginalFilename());

        Workbook workbook = null;
        if (extension.equals("xlsx")) {
            workbook = new XSSFWorkbook(file.getInputStream());
        } else if (extension.equals("xls")) {
            workbook = new HSSFWorkbook(file.getInputStream());
        }
        return workbook;
    }

}
