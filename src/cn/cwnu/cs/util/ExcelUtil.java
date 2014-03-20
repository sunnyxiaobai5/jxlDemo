package cn.cwnu.cs.util;

import java.io.File;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Set;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class ExcelUtil {
	/**
	 * 根据指定路径获取 excel 文件
	 * 
	 * @param filePath
	 *            excel 文件路径
	 * @return 要获取的 excel 文件
	 */
	public static Workbook getWorkbook(String filePath) {
		Workbook wb = null;
		try {
			wb = Workbook.getWorkbook(new File(filePath.trim()));
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wb;
	}

	/**
	 * 根据指定路径和下标获取 excel 文件中的某个只读的 Sheet
	 * 
	 * @param filePath
	 *            excel 文件路径
	 * @param sheetIndex
	 *            sheet 索引下标
	 * @return 要获取的 sheet
	 */
	public static Sheet getSheet(String filePath, int sheetIndex) {
		return ExcelUtil.getWorkbook(filePath).getSheet(sheetIndex);
	}

	/**
	 * 输出数据到 excel 文件中
	 * 
	 * @param filePath
	 *            excel 文件路径
	 */
	public static void writeExcel(String filePath) {
		WritableWorkbook wb = null;
		try {
			wb = Workbook.createWorkbook(new File(filePath));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		if (wb != null) {
			WritableSheet sheet = wb.createSheet("sheet1", 0);
			for (int i = 0; i < 10; i++) {
				for (int j = 0; j < 10; j++) {
					Label label = new Label(j, i, (i + 1) + "行" + (j + 1) + "列");
					try {
						sheet.addCell(label);
					} catch (RowsExceededException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (WriteException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
			try {
				wb.write();
				wb.close();
			} catch (WriteException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
	}

	/**
	 * 计算两列的乘积并将结果放入最后一列
	 * 
	 * @param filePath
	 *            要处理的 excel 文件地址
	 * @param sheetIndex
	 *            要处理的 sheet 下标
	 * @param rowStart
	 *            要处理的起始行下标
	 * @param col1
	 *            要处理的第一行
	 * @param col2
	 *            要处理的第二行
	 * @return excel 文件中数据组成的字符串
	 */
	public static void computeResult(String filePath, int sheetIndex,
			int rowStart, int col1, int col2) {

		// 获取读入数据的 Workbook
		Workbook rwb = ExcelUtil.getWorkbook(filePath);

		// 打开读入数据 Workbook 的副本，并指定数据写回读入数据的 Workbook
		WritableWorkbook wwb = null;
		try {
			wwb = Workbook.createWorkbook(new File(filePath), rwb);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// 获取要写入数据的 Sheet
		WritableSheet wsheet = wwb.getSheet(sheetIndex);

		// 获取写入数据的起始列
		int outputCol = wsheet.getColumns();

		try {
			// 获取读入数据的 Sheet
			Sheet sheet = rwb.getSheet(sheetIndex);
			int rowNum = sheet.getRows();
			for (int i = rowStart; i < rowNum; i++) {
				int num1 = Integer.parseInt(sheet.getRow(i)[col1].getContents()
						.trim());
				int num2 = Integer.parseInt(sheet.getRow(i)[col2].getContents()
						.trim());
				wsheet.addCell(new Number(outputCol, i, num1 * num2));
			}

			rwb.close();
			wwb.write();
			wwb.close();
		} catch (RowsExceededException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (WriteException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	/**
	 * 计算两列的乘积并将结果放入最后一列
	 * 
	 * @param rwb
	 *            要读入数据的 Workbook
	 * @param wwb
	 *            要写入数据的 Workbook
	 * @param shName
	 *            要处理的 Sheet 名称
	 * @param rowStart
	 *            要处理的起始行下标
	 * @param col1
	 *            要处理的第一列
	 * @param col2
	 *            要处理的第二列
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	public static void computeResult(Workbook rwb, WritableWorkbook wwb,
			String shName, int rowStart, int col1, int col2)
			throws RowsExceededException, WriteException {
		// 获取要读入数据的 Sheet
		Sheet sheet = rwb.getSheet(shName);

		// 获取要写入数据的 Sheet
		WritableSheet wsheet = wwb.getSheet(shName.trim());

		// 获取写入数据的起始列
		int outputCol = wsheet.getColumns();

		int rowNum = sheet.getRows();
		for (int i = rowStart; i < rowNum; i++) {
			int num1 = Integer.parseInt(sheet.getRow(i)[col1].getContents()
					.trim());
			int num2 = Integer.parseInt(sheet.getRow(i)[col2].getContents()
					.trim());
			wsheet.addCell(new Number(outputCol, i, num1 * num2));
		}

	}

	/**
	 * 根据 colBasis 列里面的值进行分类汇总，有多少個不同的值就会生成多少个 Sheet
	 * 
	 * @param rwb
	 *            要读入数据的 Workbook
	 * @param wwb
	 *            要写入数据的 Workbook
	 * @param shName
	 *            要处理的 Sheet 名称
	 * @param rowStart
	 *            要处理的起始行下标
	 * @param colBasis
	 *            处理时分类所依据列的下标
	 * @throws WriteException
	 * @throws RowsExceededException
	 * 
	 */
	public static void classification(Workbook rwb, WritableWorkbook wwb,
			String shName, int rowStart, int colBasis)
			throws RowsExceededException, WriteException {

		// 获取要读入数据的 Sheet
		Sheet rsh = rwb.getSheet(shName);

		// 获取要读入数据 Sheet 的行数
		int srn = rsh.getRows() - rowStart;

		// 获取要处理数据的总行数
		int prn = srn - rowStart;
		System.out.println(prn + "=====");

		// 获取要读入数据 Sheet 的列数
		int rcn = rsh.getColumns();

		// 获取 colBasis 列的不重复集合
		Set<Integer> set = new HashSet<Integer>();
		for (int i = rowStart; i < prn; i++) {
			System.out.println(i);
			String strValue = rsh.getCell(colBasis, i).getContents().trim();
			if (strValue == null || strValue == "") {
				break;
			}
			int value = Integer.parseInt(strValue);
			set.add(new Integer(value));
		}

		// 获取 colBasis 列形成 Set 的 Iterator ，以便根据其中的值对数据进行分类汇总
		Iterator<Integer> it = set.iterator();

		while (it.hasNext()) {
			int val = it.next().intValue();

			// 创建要写入数据的 Sheet
			int shIndex = wwb.getNumberOfSheets();
			WritableSheet writeSheet = wwb.createSheet("員工" + val, shIndex);

			for (int i = rowStart; i < prn; i++) {
				String strValue = rsh.getCell(colBasis, i).getContents().trim();
				if (strValue == null || strValue == "") {
					break;
				}
				int nextVal = Integer.parseInt(strValue);
				if (val == nextVal) {
					int writeRow = writeSheet.getRows();
					for (int j = 0; j < rcn; j++) {
						writeSheet.addCell(new Label(j, writeRow, rsh
								.getCell(j, i).getContents().trim()));
					}
				}
			}
		}
	}
	
}
