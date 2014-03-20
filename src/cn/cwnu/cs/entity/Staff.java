package cn.cwnu.cs.entity;

import java.io.File;
import java.io.IOException;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;

public class Staff {
	// excel �ļ�Ŀ¼
	private String fileDir;
	// excel �ļ������������ļ���׺��
	private String excelName;
	// excel �ļ�����·��
	private String excelRealPath;
	// Ҫ�������ݵ� Workbook
	private Workbook rwb;
	// Ҫд�����ݵ� Workbook
	private WritableWorkbook wwb;
	// Ա��������
	private Sheet shStaff;
	// ���﹤����
	private Sheet shGoods;
	// ������Ϣ��
	private Sheet shInfo;
	// Ա��д�빤����
	private WritableSheet wshStaff;
	// ����д����Ϣ��
	private WritableSheet wshGoods;
	// ����д����Ϣ��
	private WritableSheet wshInfo;

	public String getFileDir() {
		return fileDir;
	}

	public void setFileDir(String fileDir) {
		this.fileDir = fileDir;
	}

	public String getExcelName() {
		return excelName;
	}

	public void setExcelName(String excelName) {
		this.excelName = excelName;
	}

	public String getExcelRealPath() {
		return excelRealPath;
	}

	public Workbook getRwb() {
		return rwb;
	}

	public void setRwb(Workbook rwb) {
		this.rwb = rwb;
	}

	public WritableWorkbook getWwb() {
		return wwb;
	}

	public void setWwb(WritableWorkbook wwb) {
		this.wwb = wwb;
	}

	public Sheet getShStaff() {
		return shStaff;
	}

	public void setShStaff(Sheet shStaff) {
		this.shStaff = shStaff;
	}

	public Sheet getShGoods() {
		return shGoods;
	}

	public void setShGoods(Sheet shGoods) {
		this.shGoods = shGoods;
	}

	public Sheet getShInfo() {
		return shInfo;
	}

	public void setShInfo(Sheet shInfo) {
		this.shInfo = shInfo;
	}

	public void init(String shStaffName, String shGoodsName, String shInfoName) {
		// ���� excel �ļ�����·��
		this.excelRealPath = this.fileDir.trim() + this.excelName.trim()
				+ ".xls";
		try {
			this.rwb = Workbook.getWorkbook(new File(this.excelRealPath));
			this.wwb = Workbook.createWorkbook(new File(this.excelRealPath),
					this.rwb);
		} catch (BiffException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		this.shStaff = this.rwb.getSheet(shStaffName.trim());
		this.shGoods = this.rwb.getSheet(shGoodsName.trim());
		this.shInfo = this.rwb.getSheet(shInfoName.trim());
		this.wshStaff = this.wwb.getSheet(this.shStaff.getName());
		this.wshGoods = this.wwb.getSheet(this.shGoods.getName());
		this.wshInfo = this.wwb.getSheet(this.shInfo.getName());
	}

	public void close() throws IOException, WriteException {
		this.rwb.close();
		this.wwb.write();
		this.wwb.close();
	}
}
