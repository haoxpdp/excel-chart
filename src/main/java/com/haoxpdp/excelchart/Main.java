package com.haoxpdp.excelchart;

import com.haoxpdp.excelchart.bean.Position;
import com.haoxpdp.excelchart.util.DrawUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.stream.Collectors;
import java.util.stream.Stream;

public class Main {
    public static void main(String[] args) throws IOException {
        String tmpExcel = "/tmp/tmp.xlsx";

        String[] horizontalTitle = {"1", "2", "3", "4", "5"};
        String[] verticalTitle = {"a", "b"};
        Position start = new Position(1, 1);
        Position end = new Position(1 + horizontalTitle.length, 1 + verticalTitle.length);

        File file = new File(tmpExcel);
        createFileDeleteIfExists(file);
        try (Workbook wb = new XSSFWorkbook();
             OutputStream outputStream = new FileOutputStream(file)) {
            XSSFSheet sheet = (XSSFSheet) wb.createSheet();
            writeTitle(horizontalTitle, sheet);
            writeVerticalTitle(verticalTitle, sheet);
            writeData(start, end, sheet);
            drawBarChart(start, end, sheet, horizontalTitle, verticalTitle);
            wb.write(outputStream);
        }
    }

    private static void drawBarChart(Position dataStart,Position dataEnd,XSSFSheet sheet,String[] horizontalTitle,String[] verticalTitle){
        Position chartStart = new Position(0,0);
        Position chartEnd = new Position(chartStart.getX() + 6, chartStart.getY() + 4);
        String[] dataRef = {
                sheet.getSheetName()+"!$B2:$F2",
                sheet.getSheetName()+"!$B3:$F3"
        };
        DrawUtil.drawBarChart(sheet,chartStart,chartEnd,
                Stream.of(horizontalTitle).collect(Collectors.toList()),
                Stream.of(verticalTitle).collect(Collectors.toSet()),
                Stream.of(dataRef).collect(Collectors.toList()));
    }

    private static void writeData(Position start, Position end, XSSFSheet xssfSheet) {
        for (int i = start.getY(); i < end.getY(); i++) {
            Row row = getRow(i, xssfSheet);
            for (int j = start.getX(); j < end.getX(); j++) {
                Cell cell = getCell(j, row);
                cell.setCellValue(Math.random() * 10);
            }
        }
    }

    private static void writeTitle(String[] title, XSSFSheet sheet) {
        Row titleRow = getRow(0, sheet);
        Stream.of(title)
                .forEach(s -> {
                    Cell cell = getCell(Integer.valueOf(s), titleRow);
                    cell.setCellValue(s);
                });
    }

    private static void writeVerticalTitle(String[] title, XSSFSheet sheet) {
        for (int i = 0; i < title.length; i++) {
            int line = i + 1;
            Row row = getRow(line, sheet);
            Cell cell = getCell(0, row);
            cell.setCellValue(title[i]);
        }
    }

    private static Cell getCell(int index, Row row) {
        Cell cell = row.getCell(index);
        if (cell == null) cell = row.createCell(index);
        return cell;
    }

    private static Row getRow(int line, Sheet sheet) {
        Row row = sheet.getRow(line);
        if (row == null) row = sheet.createRow(line);
        return row;
    }

    private static void createFileDeleteIfExists(File file) throws IOException {
        if (file.exists()) {
            file.delete();
        } else {
            if (!file.getParentFile().exists())
                file.getParentFile().mkdirs();
        }

        file.createNewFile();

    }
}
