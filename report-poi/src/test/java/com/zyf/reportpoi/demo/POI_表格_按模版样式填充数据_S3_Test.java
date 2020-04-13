package com.zyf.reportpoi.demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

public class POI_表格_按模版样式填充数据_S3_Test {
    public static void main(String[] args) throws Exception {
        POI_表格_按模版样式填充数据_S3_Test t = new POI_表格_按模版样式填充数据_S3_Test();
        t.insertDataToTable("f:/saveFile/temp/sys_s3_template.docx", 3, false);
    }

    public void insertDataToTable(String filePath, int tableSize,
            boolean isDelTmpRow) throws Exception {
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        List<List<String>> resultList = generateTestData(4);
        insertValueToTable(doc, resultList, tableSize, isDelTmpRow);
        saveDocument(doc, "f:/saveFile/temp/sys_" + System.currentTimeMillis()
                + ".docx");
    }

    /**
     * @Description: 按模版行样式填充数据,暂未实现特殊样式填充(如列合并)，只能用于普通样式(如段落间距 缩进 字体 对齐)
     * @param resultList  填充数据
     * @param tableRowSize  模版表格行数 取第一个行数相等列数相等的表格填充
     * @param isDelTmpRow 是否删除模版行
     */
    // TODO 数据行插到模版行下面，没有实现指定位置插入
    public void insertValueToTable(XWPFDocument doc,
            List<List<String>> resultList, int tableRowSize, boolean isDelTmpRow)
            throws Exception {
        Iterator<XWPFTable> iterator = doc.getTablesIterator();
        XWPFTable table = null;
        List<XWPFTableRow> rows = null;
        List<XWPFTableCell> cells = null;
        List<XWPFTableCell> tmpCells = null;// 模版列
        XWPFTableRow tmpRow = null;// 匹配用
        XWPFTableCell tmpCell = null;// 匹配用
        boolean flag = false;// 是否找到表格
        while (iterator.hasNext()) {
            table = iterator.next();
            rows = table.getRows();
            if (rows.size() == tableRowSize) {
                tmpRow = rows.get(tableRowSize - 1);
                cells = tmpRow.getTableCells();
                if (cells.size() == resultList.get(0).size()) {
                    flag = true;
                    break;
                }
            }
        }
        if (!flag) {
            return;
        }
        tmpCells = tmpRow.getTableCells();
        for (int i = 0, len = resultList.size(); i < len; i++) {
            XWPFTableRow row = table.createRow();
            row.setHeight(tmpRow.getHeight());
            List<String> list = resultList.get(i);
            cells = row.getTableCells();
            // 插入的行会填充与表格第一行相同的列数
            for (int k = 0, klen = cells.size(); k < klen; k++) {
                tmpCell = tmpCells.get(k);
                XWPFTableCell cell = cells.get(k);
                setCellText(tmpCell, cell, list.get(k));
            }
            // 继续写剩余的列
            for (int j = cells.size(), jlen = list.size(); j < jlen; j++) {
                tmpCell = tmpCells.get(j);
                XWPFTableCell cell = row.addNewTableCell();
                setCellText(tmpCell, cell, list.get(j));
            }
        }
        // 删除模版行
        if (isDelTmpRow) {
            table.removeRow(tableRowSize - 1);
        }
    }

    public void setCellText(XWPFTableCell tmpCell, XWPFTableCell cell,
            String text) throws Exception {
        CTTc cttc2 = tmpCell.getCTTc();
        CTTcPr ctPr2 = cttc2.getTcPr();

        CTTc cttc = cell.getCTTc();
        CTTcPr ctPr = cttc.addNewTcPr();
        cell.setColor(tmpCell.getColor());
        cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
        if (ctPr2.getTcW() != null) {
            ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
        }
        if (ctPr2.getVAlign() != null) {
            ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
        }
        if (cttc2.getPList().size() > 0) {
            CTP ctp = cttc2.getPList().get(0);
            if (ctp.getPPr() != null) {
                if (ctp.getPPr().getJc() != null) {
                    cttc.getPList().get(0).addNewPPr().addNewJc().setVal(
                            ctp.getPPr().getJc().getVal());
                }
            }
        }

        if (ctPr2.getTcBorders() != null) {
            ctPr.setTcBorders(ctPr2.getTcBorders());
        }

        XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
        XWPFParagraph cellP = cell.getParagraphs().get(0);
        XWPFRun tmpR = null;
        if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
            tmpR = tmpP.getRuns().get(0);
        }
        XWPFRun cellR = cellP.createRun();
        cellR.setText(text);
        // 复制字体信息
        if (tmpR != null) {
            cellR.setBold(tmpR.isBold());
            cellR.setItalic(tmpR.isItalic());
            cellR.setStrike(tmpR.isStrike());
            cellR.setUnderline(tmpR.getUnderline());
            cellR.setColor(tmpR.getColor());
            cellR.setTextPosition(tmpR.getTextPosition());
            if (tmpR.getFontSize() != -1) {
                cellR.setFontSize(tmpR.getFontSize());
            }
            if (tmpR.getFontFamily() != null) {
                cellR.setFontFamily(tmpR.getFontFamily());
            }
            if (tmpR.getCTR() != null) {
                if (tmpR.getCTR().isSetRPr()) {
                    CTRPr tmpRPr = tmpR.getCTR().getRPr();
                    if (tmpRPr.isSetRFonts()) {
                        CTFonts tmpFonts = tmpRPr.getRFonts();
                        CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR
                                .getCTR().getRPr() : cellR.getCTR().addNewRPr();
                        CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr
                                .getRFonts() : cellRPr.addNewRFonts();
                        cellFonts.setAscii(tmpFonts.getAscii());
                        cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
                        cellFonts.setCs(tmpFonts.getCs());
                        cellFonts.setCstheme(tmpFonts.getCstheme());
                        cellFonts.setEastAsia(tmpFonts.getEastAsia());
                        cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
                        cellFonts.setHAnsi(tmpFonts.getHAnsi());
                        cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
                    }
                }
            }
        }
        // 复制段落信息
        cellP.setAlignment(tmpP.getAlignment());
        cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
        cellP.setBorderBetween(tmpP.getBorderBetween());
        cellP.setBorderBottom(tmpP.getBorderBottom());
        cellP.setBorderLeft(tmpP.getBorderLeft());
        cellP.setBorderRight(tmpP.getBorderRight());
        cellP.setBorderTop(tmpP.getBorderTop());
        cellP.setPageBreak(tmpP.isPageBreak());
        if (tmpP.getCTP() != null) {
            if (tmpP.getCTP().getPPr() != null) {
                CTPPr tmpPPr = tmpP.getCTP().getPPr();
                CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP
                        .getCTP().getPPr() : cellP.getCTP().addNewPPr();
                // 复制段落间距信息
                CTSpacing tmpSpacing = tmpPPr.getSpacing();
                if (tmpSpacing != null) {
                    CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr
                            .getSpacing()
                            : cellPPr.addNewSpacing();
                    if (tmpSpacing.getAfter() != null) {
                        cellSpacing.setAfter(tmpSpacing.getAfter());
                    }
                    if (tmpSpacing.getAfterAutospacing() != null) {
                        cellSpacing.setAfterAutospacing(tmpSpacing
                                .getAfterAutospacing());
                    }
                    if (tmpSpacing.getAfterLines() != null) {
                        cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
                    }
                    if (tmpSpacing.getBefore() != null) {
                        cellSpacing.setBefore(tmpSpacing.getBefore());
                    }
                    if (tmpSpacing.getBeforeAutospacing() != null) {
                        cellSpacing.setBeforeAutospacing(tmpSpacing
                                .getBeforeAutospacing());
                    }
                    if (tmpSpacing.getBeforeLines() != null) {
                        cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
                    }
                    if (tmpSpacing.getLine() != null) {
                        cellSpacing.setLine(tmpSpacing.getLine());
                    }
                    if (tmpSpacing.getLineRule() != null) {
                        cellSpacing.setLineRule(tmpSpacing.getLineRule());
                    }
                }
                // 复制段落缩进信息
                CTInd tmpInd = tmpPPr.getInd();
                if (tmpInd != null) {
                    CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd()
                            : cellPPr.addNewInd();
                    if (tmpInd.getFirstLine() != null) {
                        cellInd.setFirstLine(tmpInd.getFirstLine());
                    }
                    if (tmpInd.getFirstLineChars() != null) {
                        cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
                    }
                    if (tmpInd.getHanging() != null) {
                        cellInd.setHanging(tmpInd.getHanging());
                    }
                    if (tmpInd.getHangingChars() != null) {
                        cellInd.setHangingChars(tmpInd.getHangingChars());
                    }
                    if (tmpInd.getLeft() != null) {
                        cellInd.setLeft(tmpInd.getLeft());
                    }
                    if (tmpInd.getLeftChars() != null) {
                        cellInd.setLeftChars(tmpInd.getLeftChars());
                    }
                    if (tmpInd.getRight() != null) {
                        cellInd.setRight(tmpInd.getRight());
                    }
                    if (tmpInd.getRightChars() != null) {
                        cellInd.setRightChars(tmpInd.getRightChars());
                    }
                }
            }
        }
    }

    public void saveDocument(XWPFDocument document, String savePath)
            throws Exception {
        FileOutputStream fos = new FileOutputStream(savePath);
        document.write(fos);
        fos.close();
    }

    // 生成测试数据
    public List<List<String>> generateTestData(int num) {
        List<List<String>> resultList = new ArrayList<List<String>>();
        for (int i = 1; i <= num; i++) {
            List<String> list = new ArrayList<String>();
            list.add("" + i);
            list.add("测试_" + i);
            list.add("测试2_" + i);
            list.add("测试3_" + i);
            list.add("测试4_" + i);
            resultList.add(list);
        }
        return resultList;
    }
}