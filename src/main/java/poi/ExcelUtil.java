package poi;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



/**
 * Excel文件处理工具
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
		Workbook workbook = new XSSFWorkbook();
		// 生成名为"first"的工作表，参数0表示这是第一页
		Sheet sheet = workbook.createSheet("first");
		//填充文本  单元格位置是[(0,0)=(A,1)|(1,2)=(B,3)]
		for (int i=0; i<dataList.size(); i++) {
			Row row = sheet.createRow(i);
			List<Object> list = dataList.get(i);
			for (int j=0; j<list.size(); j++) {
				row.createCell(j).setCellValue(list.get(j)==null ? "" :list.get(j).toString());
			}
		}
		workbook.write(outputStream);
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
		Workbook workbook = new XSSFWorkbook(inputStream);
		// 获取第一页工作表
		Sheet sheet = workbook.getSheetAt(0); 
		List<List<Object>> dataList = new ArrayList<>(sheet.getLastRowNum());
		List<Object> rowList = null;
		for(int i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++) {
			Row row = sheet.getRow(i);
			rowList = new ArrayList<>(row.getLastCellNum());
			for(int j=row.getFirstCellNum(); j<=row.getLastCellNum(); j++) {
				Cell cell = row.getCell(j);
				rowList.add(cell.getStringCellValue());
			}
			dataList.add(rowList);
		}
		workbook.close();
		return dataList;
	}
	
}
