package com.zyf.reportpoi.demo;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.TextAlignment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;

import com.deepoove.poi.policy.DynamicTableRenderPolicy;

/**
 * 表格标签插件 
 * <br>支持一行表头</br>
 * <br>支持交替继承数据行背景色、行高</br>
 * <br>支持继承数据行的第一行的水平位置</br>
 */
public class TableRenderPolicy2 extends DynamicTableRenderPolicy {

	// 支持的模板表格标题行数
	private final int TITLE_ROWS_SIZE = 1;

	/**
	 * 绑定数据到表格
	 * 
	 * @param table
	 *            表格
	 * @param data
	 *            数据
	 */
	@SuppressWarnings({"unchecked","unused"})
	@Override
	public void render(XWPFTable table, Object data) {
		if (null == data) {
			return;
		}
		// 表格数据强制转换
		List<List<String>> rowList = new ArrayList<>();
		if (data instanceof List) {
			rowList = (List<List<String>>) data;
		} else {
			return;
		}

		// 判断要填充的数据行数是否超出现有的表格行数，超出的话，提前插入新行（复制原来行的背景色）
		if (rowList.size() > table.getRows().size()) {
			// 模板表格总行数
			int templateRowsSize = table.getRows().size();
			// 数据行数
			int dataRowsSize = rowList.size();
			// 数据列数
			int dataColumnSize = rowList.get(0).size();
			// 模板表格的背景色
			List<String> color = new ArrayList<>();
			// 模板表格的行高
			List<Integer> hight = new ArrayList<>();
			// 模板表格水平位置
			List<ParagraphAlignment> paragraphAlignment = new ArrayList<>();
			// 模板表格垂直位置
			List<TextAlignment> textAlignment = new ArrayList<>();
			// 模板表格行属性
			List<CTTrPr> cTTrPrs = new ArrayList<>();
			List<XWPFTableCell> cellList = new ArrayList<>();
			for (int i = 0; i < templateRowsSize; i++) {
				if (TITLE_ROWS_SIZE < i + 1) { // 不取标题行样式
					color.add(table.getRows().get(i).getCell(0).getColor());
					hight.add(table.getRows().get(i).getHeight());
				}
				if (TITLE_ROWS_SIZE == i) {// 只取标题后第一行样式
					for (int j = 0; j < table.getRows().size(); j++) {
						XWPFTableCell cell = table.getRows().get(i).getCell(j);
						XWPFParagraph cellParagraph = cell.getParagraphs().get(0);
						paragraphAlignment.add(cellParagraph.getAlignment());
						textAlignment.add(cellParagraph.getVerticalAlignment());
						/////
						cellList = table.getRows().get(i).getTableCells();
						//////
						 XWPFParagraph tmpP = cell.getParagraphs().get(0);
					}
					/////////
//					table.getRows().get(i).getCtRow().setTrPr(table.getRows().get(i).getCtRow().getTrPr());
					CTTrPr cTTrPr= table.getRows().get(i).getCtRow().getTrPr();
					cTTrPrs.add(cTTrPr);
					
				}
			}
			// 补充表格行数
			for (int i = templateRowsSize; i < dataRowsSize + TITLE_ROWS_SIZE; i++) {
				// 新建行
				XWPFTableRow newRow = table.insertNewTableRow(i);
				
				if (hight.size() > 0) {
					// 交替设置行高
//					newRow.setHeight(hight.get((i - TITLE_ROWS_SIZE) % (hight.size())));
				}
				if (null != newRow) {
					// 遍历获取单元格
					for (int j = 0; j < dataColumnSize; j++) {
						// 新建单元格
						XWPFTableCell newCell = newRow.createCell();
						newRow.getCtRow().setTrPr(cTTrPrs.get(0));
						if (color.size() > 0) {
							// 交替设置单元格颜色
//							newCell.setColor(color.get((i - TITLE_ROWS_SIZE) % (color.size())));
						}
						if (paragraphAlignment.size() > 0) {
//							newCell.getParagraphs().get(0).setAlignment(paragraphAlignment.get(j));
						}
						if (textAlignment.size() > 0) {
//							newCell.getParagraphs().get(0).setVerticalAlignment(textAlignment.get(j));
						}
						
						XWPFTableCell sourceCell = cellList.get(j);
						// 列属性
						newCell.getCTTc().setTcPr(sourceCell.getCTTc().getTcPr());
						// 段落属性
						if (sourceCell.getParagraphs() != null && sourceCell.getParagraphs().size() > 0) {
							newCell.getParagraphs().get(0).getCTP().setPPr(sourceCell.getParagraphs().get(0).getCTP().getPPr());
							if (sourceCell.getParagraphs().get(0).getRuns() != null
									&& sourceCell.getParagraphs().get(0).getRuns().size() > 0) {
								XWPFRun cellR = newCell.getParagraphs().get(0).createRun();
								cellR.setText(sourceCell.getText());
								cellR.setBold(sourceCell.getParagraphs().get(0).getRuns().get(0).isBold());
							} else {
								newCell.setText(sourceCell.getText());
							}
						} else {
							newCell.setText(sourceCell.getText());
						}
					    
					}
				}
			}
		}
		// 填写单元格数据
		for (int i = 0; i < rowList.size(); i++) {
			List<String> rowSet = rowList.get(i);
			for (int j = 0; j < rowSet.size(); j++) {
				table.getRow(i + TITLE_ROWS_SIZE).getCell(j).setText(rowSet.get(j));
			}
		}
	}
}