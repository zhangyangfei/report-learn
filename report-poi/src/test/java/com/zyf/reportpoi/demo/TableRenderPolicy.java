package com.zyf.reportpoi.demo;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

import com.deepoove.poi.policy.DynamicTableRenderPolicy;

public class TableRenderPolicy extends DynamicTableRenderPolicy {

	@Override
	public void render(XWPFTable table, Object data) {
		// 数据强制转换一下
		List<List<String>> rowList = new ArrayList<>();
		if (data instanceof Set) {
			rowList = (List<List<String>>) data;
		} else {
			rowList = (List<List<String>>) data;
		}
		// 判断要填充的数据行数是否超出现有的表格行数 超出的话 进行插入新行提前插入
		if (null != rowList) {
			if (rowList.size() > table.getRows().size()) {
				List<String> color = new ArrayList<>();
				if (table.getRows().size() == 3) {
					color.add(table.getRows().get(1).getCell(0).getColor());
					color.add(table.getRows().get(2).getCell(0).getColor());
				}

				for (int i = table.getRows().size() - 1; i < rowList.size(); i++) {
					XWPFTableRow insertNewTableRow = null;

					insertNewTableRow = table.insertNewTableRow(i - 1);

					if (null != insertNewTableRow) {
						for (int j = 0; j < rowList.get(i).size(); j++) {
							XWPFTableCell cell = insertNewTableRow.createCell();
							if (color.size() > 0) {
								String col = color.get(0);
								if (i % 2 == 1) {
									col = color.get(1);
								}
								cell.setColor(col);
							}
						}
					}
				}
			}

			// 循环数据 将数据一次放入每个单元格
			for (int i = 0; i < rowList.size(); i++) {
				List<String> rowSet = rowList.get(i);
				for (int j = 0; j < rowSet.size(); j++) {
					table.getRow(i + 1).getCell(j).setText(rowSet.get(j));
				}
			}
		}
	}
}