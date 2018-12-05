package com.haoxpdp.excelchart.util;

import com.haoxpdp.excelchart.bean.Position;
import org.apache.poi.ss.usermodel.Chart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.charts.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;

import java.util.Iterator;
import java.util.List;
import java.util.Set;

public class DrawUtil {
    /*
     * @param params params.get(0) 为图表位置 0123为起始横坐标，起始纵坐标，终点横，终点纵
     *
     * @param params params.get(1) 为x轴取值范围，其他为折线取值范围。起始行号，终止行号， 起始列号，终止列号（折线名称坐标
     * ）。
     */
    private static void drawLineChart(XSSFSheet sheet, List<int[]> params) {

        Drawing drawing = sheet.createDrawingPatriarch();
        // 设置位置 起始横坐标，起始纵坐标，终点横，终点纵
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, params.get(0)[0], params.get(0)[1], params.get(0)[2],
                params.get(0)[3]);
        Chart chart = drawing.createChart(anchor);

        // 创建标题

        // 创建图形注释的位置
        ChartLegend legend = chart.getOrCreateLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        // 创建绘图的类型 LineCahrtData 折线图
        LineChartData chartData = chart.getChartDataFactory().createLineChartData();
        // 设置横坐标
        ChartAxis bottomAxis = chart.getChartAxisFactory().createCategoryAxis(AxisPosition.BOTTOM);
        bottomAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // 设置纵坐标
        ValueAxis leftAxis = chart.getChartAxisFactory().createValueAxis(AxisPosition.LEFT);

        leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
        // 从Excel中获取数据.CellRangeAddress-->params: 起始行号，终止行号， 起始列号，终止列号
        // 数据类别（x轴）
        ChartDataSource xAxis = DataSources.fromStringCellRange(sheet,
                new CellRangeAddress(params.get(1)[0], params.get(1)[1], params.get(1)[2], params.get(1)[3]));
        for (int i = 2, len = params.size(); i < len; i++) {
            // 数据源
            ChartDataSource dataAxis = DataSources.fromNumericCellRange(sheet,
                    new CellRangeAddress(params.get(i)[0], params.get(i)[1], params.get(i)[2], params.get(i)[3]));

            LineChartSeries chartSeries = chartData.addSeries(xAxis, dataAxis);
            // 给每条折线创建名字
            chartSeries.setTitle(sheet.getRow(params.get(i)[4]).getCell(params.get(i)[5]).getStringCellValue());
        }

        setChartTitle((XSSFChart) chart, sheet.getSheetName());
        // 开始绘制折线图
        chart.plot(chartData, bottomAxis, leftAxis);
    }

    public static void drawCTLineChart(XSSFSheet sheet, Position start, Position end, List<String> xString, Set<String> serTxName,
                                       List<String> dataRef) {

        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, start.getX(), start.getY(), end.getX(), end.getY());

        Chart chart = drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart) chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTLineChart ctLineChart = ctPlotArea.addNewLineChart();

        ctLineChart.addNewVaryColors().setVal(true);

        // telling the Chart that it has axis and giving them Ids
        setAxIds(ctLineChart);
        // set cat axis
        setCatAx(ctPlotArea);
        // set val axis
        setValAx(ctPlotArea);
        // legend
        setLegend(ctChart);
        // set data lable
        setDataLabel(ctLineChart);
        // set chart title
        setChartTitle((XSSFChart) chart, sheet.getSheetName());

        paddingData(ctLineChart, xString, serTxName, dataRef);
    }

    /*
     * @param position 图表坐标 起始行，起始列，终点行，重点列
     *
     * @param xString 横坐标
     *
     * @param serTxName 图形示例
     *
     * @param dataRef 柱状图数据范围 ： sheetName!$A$1:$A$12
     */
    public static void drawBarChart(XSSFSheet sheet, Position start, Position end, List<String> xString, Set<String> serTxName,
                                    List<String> dataRef) {

        Drawing drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, start.getX(), start.getY(), end.getX(), end.getY());

        Chart chart = drawing.createChart(anchor);

        CTChart ctChart = ((XSSFChart) chart).getCTChart();
        CTPlotArea ctPlotArea = ctChart.getPlotArea();
        CTBarChart ctBarChart = ctPlotArea.addNewBarChart();

        ctBarChart.addNewVaryColors().setVal(false);
        ctBarChart.addNewBarDir().setVal(STBarDir.COL);

        // telling the Chart that it has axis and giving them Ids
        setAxIds(ctBarChart);
        // set cat axis
        setCatAx(ctPlotArea);
        // set val axis
        setValAx(ctPlotArea);
        // add legend and set legend position
        setLegend(ctChart);
        // set data lable
        setDataLabel(ctBarChart);
        // set chart title
        setChartTitle((XSSFChart) chart, sheet.getSheetName());
        // padding data to chart
        paddingData(ctBarChart, xString, serTxName, dataRef);
    }

    private static void paddingData(CTBarChart ctBarChart, List<String> xString, Set<String> serTxName,
                                    List<String> dataRef) {
        Iterator<String> iterator = serTxName.iterator();
        for (int r = 0, len = dataRef.size(); r < len && iterator.hasNext(); r++) {
            CTBarSer ctBarSer = ctBarChart.addNewSer();

            ctBarSer.addNewIdx().setVal(r);
            // set legend value
            setLegend(iterator.next(), ctBarSer.addNewTx());
            // cat ax value
            setChartCatAxLabel(ctBarSer.addNewCat(), xString);
            // value range
            ctBarSer.addNewVal().addNewNumRef().setF(dataRef.get(r));
            // add border to chart
            ctBarSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});
        }
    }

    private static void setLegend(String str, CTSerTx ctSerTx) {
        if (str.contains("$"))
            // set legend by str ref
            ctSerTx.addNewStrRef().setF(str);
        else
            // set legend by str
            ctSerTx.setV(str);
    }

    private static void paddingData(CTLineChart ctLineChart, List<String> xString, Set<String> serTxName,
                                    List<String> dataRef) {
        Iterator<String> iterator = serTxName.iterator();
        for (int r = 0, len = dataRef.size(); r < len && iterator.hasNext(); r++) {
            CTLineSer ctLineSer = ctLineChart.addNewSer();
            ctLineSer.addNewIdx().setVal(r);
            setLegend(iterator.next(), ctLineSer.addNewTx());
            setChartCatAxLabel(ctLineSer.addNewCat(), xString);
            ctLineSer.addNewVal().addNewNumRef().setF(dataRef.get(r));
            ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSrgbClr().setVal(new byte[]{0, 0, 0});
        }
    }

    private static void setChartCatAxLabel(CTAxDataSource cttAxDataSource, List<String> xString) {
        if (xString.size() == 1) {
            cttAxDataSource.addNewStrRef().setF(xString.get(0));
        } else {
            CTStrData ctStrData = cttAxDataSource.addNewStrLit();
            for (int m = 0, xlen = xString.size(); m < xlen; m++) {
                CTStrVal ctStrVal = ctStrData.addNewPt();
                ctStrVal.setIdx((long) m);
                ctStrVal.setV(xString.get(m));
            }
        }
    }

    private static void setDataLabel(CTBarChart ctBarChart) {
        setDLShowOpts(ctBarChart.addNewDLbls());
    }

    private static void setDataLabel(CTLineChart ctLineChart) {
        CTDLbls dlbls = ctLineChart.addNewDLbls();
        setDLShowOpts(dlbls);
        setDLPosition(dlbls, null);
    }

    private static void setDLPosition(CTDLbls dlbls, STDLblPos.Enum e) {
        if (e == null)
            dlbls.addNewDLblPos().setVal(STDLblPos.T);
        else
            dlbls.addNewDLblPos().setVal(e);
    }

    private static void setDLShowOpts(CTDLbls dlbls) {
        // 添加图形示例的字符串
        dlbls.addNewShowSerName().setVal(false);
        // 添加x轴的坐标字符串
        dlbls.addNewShowCatName().setVal(false);
        // 添加图形示例的图片
        dlbls.addNewShowLegendKey().setVal(false);
        // 添加x对应y的值---全设置成false 就没什么用处了
        // dlbls.addNewShowVal().setVal(false);
    }

    private static void setAxIds(CTBarChart ctBarChart) {
        ctBarChart.addNewAxId().setVal(123456);
        ctBarChart.addNewAxId().setVal(123457);
    }

    private static void setAxIds(CTLineChart ctLineChart) {
        ctLineChart.addNewAxId().setVal(123456);
        ctLineChart.addNewAxId().setVal(123457);
    }

    private static void setLegend(CTChart ctChart) {
        CTLegend ctLegend = ctChart.addNewLegend();
        ctLegend.addNewLegendPos().setVal(STLegendPos.B);
        ctLegend.addNewOverlay().setVal(false);
    }

    private static void setCatAx(CTPlotArea ctPlotArea) {
        CTCatAx ctCatAx = ctPlotArea.addNewCatAx();
        ctCatAx.addNewAxId().setVal(123456); // id of the cat axis
        CTScaling ctScaling = ctCatAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        ctCatAx.addNewDelete().setVal(false);
        ctCatAx.addNewAxPos().setVal(STAxPos.B);
        ctCatAx.addNewCrossAx().setVal(123457); // id of the val axis
        ctCatAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);

    }

    // 不要y轴的标签，或者y轴尽可能的窄一些
    private static void setValAx(CTPlotArea ctPlotArea) {
        CTValAx ctValAx = ctPlotArea.addNewValAx();
        ctValAx.addNewAxId().setVal(123457); // id of the val axis
        CTScaling ctScaling = ctValAx.addNewScaling();
        ctScaling.addNewOrientation().setVal(STOrientation.MIN_MAX);
        // 不现实y轴
        ctValAx.addNewDelete().setVal(true);
        ctValAx.addNewAxPos().setVal(STAxPos.L);
        ctValAx.addNewCrossAx().setVal(123456); // id of the cat axis
        ctValAx.addNewTickLblPos().setVal(STTickLblPos.NEXT_TO);
    }

    // 图标标题
    private static void setChartTitle(XSSFChart xchart, String titleStr) {
        CTChart ctChart = xchart.getCTChart();

        CTTitle title = CTTitle.Factory.newInstance();
        CTTx cttx = title.addNewTx();
        CTStrData sd = CTStrData.Factory.newInstance();
        CTStrVal str = sd.addNewPt();
        str.setIdx(123459);
        str.setV(titleStr);
        cttx.addNewStrRef().setStrCache(sd);

        ctChart.setTitle(title);
    }

}
