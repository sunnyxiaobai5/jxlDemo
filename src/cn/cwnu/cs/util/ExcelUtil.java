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
	 * ����ָ��·����ȡ excel �ļ�
	 * 
	 * @param filePath
	 *            excel �ļ�·��
	 * @return Ҫ��ȡ�� excel �ļ�
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
	 * ����ָ��·�����±��ȡ excel �ļ��е�ĳ��ֻ���� Sheet
	 * 
	 * @param filePath
	 *            excel �ļ�·��
	 * @param sheetIndex
	 *            sheet �����±�
	 * @return Ҫ��ȡ�� sheet
	 */
	public static Sheet getSheet(String filePath, int sheetIndex) {
		return ExcelUtil.getWorkbook(filePath).getSheet(sheetIndex);
	}

	/**
	 * ������ݵ� excel �ļ���
	 * 
	 * @param filePath
	 *            excel �ļ�·��
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
					Label label = new Label(j, i, (i + 1) + "��" + (j + 1) + "��");
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
	 * �������еĳ˻���������������һ��
	 * 
	 * @param filePath
	 *            Ҫ����� excel �ļ���ַ
	 * @param sheetIndex
	 *            Ҫ����� sheet �±�
	 * @param rowStart
	 *            Ҫ�������ʼ���±�
	 * @param col1
	 *            Ҫ����ĵ�һ��
	 * @param col2
	 *            Ҫ����ĵڶ���
	 * @return excel �ļ���������ɵ��ַ���
	 */
	public static void computeResult(String filePath, int sheetIndex,
			int rowStart, int col1, int col2) {

		// ��ȡ�������ݵ� Workbook
		Workbook rwb = ExcelUtil.getWorkbook(filePath);

		// �򿪶������� Workbook �ĸ�������ָ������д�ض������ݵ� Workbook
		WritableWorkbook wwb = null;
		try {
			wwb = Workbook.createWorkbook(new File(filePath), rwb);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		// ��ȡҪд�����ݵ� Sheet
		WritableSheet wsheet = wwb.getSheet(sheetIndex);

		// ��ȡд�����ݵ���ʼ��
		int outputCol = wsheet.getColumns();

		try {
			// ��ȡ�������ݵ� Sheet
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
	 * �������еĳ˻���������������һ��
	 * 
	 * @param rwb
	 *            Ҫ�������ݵ� Workbook
	 * @param wwb
	 *            Ҫд�����ݵ� Workbook
	 * @param shName
	 *            Ҫ����� Sheet ����
	 * @param rowStart
	 *            Ҫ�������ʼ���±�
	 * @param col1
	 *            Ҫ����ĵ�һ��
	 * @param col2
	 *            Ҫ����ĵڶ���
	 * @throws WriteException
	 * @throws RowsExceededException
	 */
	public static void computeResult(Workbook rwb, WritableWorkbook wwb,
			String shName, int rowStart, int col1, int col2)
			throws RowsExceededException, WriteException {
		// ��ȡҪ�������ݵ� Sheet
		Sheet sheet = rwb.getSheet(shName);

		// ��ȡҪд�����ݵ� Sheet
		WritableSheet wsheet = wwb.getSheet(shName.trim());

		// ��ȡд�����ݵ���ʼ��
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
	 * ���� colBasis �������ֵ���з�����ܣ��ж��ق���ͬ��ֵ�ͻ����ɶ��ٸ� Sheet
	 * 
	 * @param rwb
	 *            Ҫ�������ݵ� Workbook
	 * @param wwb
	 *            Ҫд�����ݵ� Workbook
	 * @param shName
	 *            Ҫ����� Sheet ����
	 * @param rowStart
	 *            Ҫ�������ʼ���±�
	 * @param colBasis
	 *            ����ʱ�����������е��±�
	 * @throws WriteException
	 * @throws RowsExceededException
	 * 
	 */
	public static void classification(Workbook rwb, WritableWorkbook wwb,
			String shName, int rowStart, int colBasis)
			throws RowsExceededException, WriteException {

		// ��ȡҪ�������ݵ� Sheet
		Sheet rsh = rwb.getSheet(shName);

		// ��ȡҪ�������� Sheet ������
		int srn = rsh.getRows() - rowStart;

		// ��ȡҪ�������ݵ�������
		int prn = srn - rowStart;
		System.out.println(prn + "=====");

		// ��ȡҪ�������� Sheet ������
		int rcn = rsh.getColumns();

		// ��ȡ colBasis �еĲ��ظ�����
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

		// ��ȡ colBasis ���γ� Set �� Iterator ���Ա�������е�ֵ�����ݽ��з������
		Iterator<Integer> it = set.iterator();

		while (it.hasNext()) {
			int val = it.next().intValue();

			// ����Ҫд�����ݵ� Sheet
			int shIndex = wwb.getNumberOfSheets();
			WritableSheet writeSheet = wwb.createSheet("�T��" + val, shIndex);

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
