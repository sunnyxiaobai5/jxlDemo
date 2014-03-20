package cn.cwnu.cs.client;

import java.io.IOException;

import jxl.read.biff.BiffException;
import jxl.write.WriteException;
import cn.cwnu.cs.entity.Staff;
import cn.cwnu.cs.util.ExcelUtil;

public class Client {
	public static void main(String[] args) throws BiffException, IOException,
			WriteException {
		Staff staff = new Staff();
		staff.setFileDir("D:/");
		staff.setExcelName("Book3");
		staff.init("Sheet1", "Sheet2", "Sheet3");
		// ExcelUtil.computeResult(staff.getRwb(), staff.getWwb(), "工作信息", 0, 2,
		// 3);
		ExcelUtil
				.classification(staff.getRwb(), staff.getWwb(), "Sheet1",3, 1);

		staff.close();
	}
}
