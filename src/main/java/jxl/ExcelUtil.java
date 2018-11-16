package jxl;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;


/**
 * Excel文件处理工具
 * 只支持xls格式文件
 * @author YangJie
 * @createTime 2015年4月7日 上午11:22:52
 */
public class ExcelUtil {
	
	/**
	 * 导出
	 * 默认只有一页 名为first
	 * @param outputStream
	 * @param dataList
	 * @return
	 * @throws Exception 
	 */
	public static OutputStream write(OutputStream outputStream, List<List<Object>> dataList) throws Exception {
		// 创建工作簿(可读可写)
		WritableWorkbook workbook = Workbook.createWorkbook(outputStream);
		// 生成名为"first"的工作表，参数0表示这是第一页
		WritableSheet sheet = workbook.createSheet("first", 0);
		//填充文本  单元格位置是[(0,0)=(A,1)|(1,2)=(B,3)]
		for (int i=0; i<dataList.size(); i++) {
			List<Object> list = dataList.get(i);
			for (int j=0; j<list.size(); j++) {
				sheet.addCell(new Label(j, i, list.get(j)==null ? "" :list.get(j).toString()));
			}
		}
		workbook.write();
		workbook.close();
		return outputStream;
	}
	
	/**
	 * 导入
	 * 默认只读取第一页
	 * @param inputStream
	 * @return
	 * @throws Exception
	 */
	public static List<List<Object>> read(InputStream inputStream) throws Exception {
		// 创建工作簿(只读)
		Workbook workbook = Workbook.getWorkbook(inputStream);
		// 获取第一页工作表
		Sheet sheet = workbook.getSheet(0); 
		int rows = sheet.getRows(); // 总行数
		int columns = sheet.getColumns(); // 总列数
		List<List<Object>> dataList = new ArrayList<>(rows);
		List<Object> rowList = null;
		for(int i=0; i<rows; i++) {
			rowList = new ArrayList<>(columns);
			for(int j=0; j<columns; j++) {
				Cell cell = sheet.getCell(j, i);
				rowList.add(cell.getContents());
			}
			dataList.add(rowList);
		}
		workbook.close();
		return dataList;
	}
	
}
