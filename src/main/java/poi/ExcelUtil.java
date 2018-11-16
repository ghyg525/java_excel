package poi;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Excel文件处理工具
 * @author YangJie [2018年11月16日]
 */
public class ExcelUtil {
	
	
	/**
	 * 读取
	 * 自动识别文件类型
	 * @param inputStream 文件输入流
	 * @return
	 * @throws Exception
	 */
	public static List<List<List<Object>>> read(InputStream inputStream) throws Exception {
		List<List<List<Object>>> list = new ArrayList<>();
		Workbook workbook = WorkbookFactory.create(inputStream);
		for(int i=0; i<workbook.getNumberOfSheets(); i++) {
			list.add(readSheet(workbook.getSheetAt(i)));
		}
		workbook.close();
		return list;
	}
	
	/**
	 * 读取
	 * 自动识别文件类型
	 * @param inputStream 文件输入流
	 * @param sheetIndex 指定sheet下标
	 * @return
	 * @throws Exception
	 */
	public static List<List<Object>> read(InputStream inputStream, int sheetIndex) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		List<List<Object>> list = readSheet(workbook.getSheetAt(sheetIndex));
		workbook.close();
		return list;
	}
	
	/**
	 * 读取 sheet
	 * @param sheet
	 * @return
	 */
	private static List<List<Object>> readSheet(Sheet sheet) {
		List<List<Object>> sheetList = new ArrayList<>(sheet.getLastRowNum());
		for(int i=sheet.getFirstRowNum(); i<=sheet.getLastRowNum(); i++) {
			sheetList.add(readRow(sheet.getRow(i)));
		}
		return sheetList;
	}
	
	/**
	 * 读取 行
	 * @param row
	 * @return
	 */
	private static List<Object> readRow(Row row) {
		List<Object> rowList = new ArrayList<>(row.getLastCellNum());
		for(int i=row.getFirstCellNum(); i<=row.getLastCellNum(); i++) {
			rowList.add(readCell(row.getCell(i)));
		}
		return rowList;
	}
	
	/**
	 * 读取 单元格
	 * @param cell
	 * @return
	 */
	private static Object readCell(Cell cell) {
		if(Objects.isNull(cell)) {
			return null;
		}
		switch (cell.getCellTypeEnum()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			return cell.getNumericCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue();
		default:
			return null;
		}
	}
	
	/**
	 * 读取图片
	 * @param inputStream
	 * @return
	 * @throws Exception
	 */
	public static List<ReadImageBean> readImage(InputStream inputStream) throws Exception {
		List<ReadImageBean> List = new ArrayList<>();
		Workbook workbook = WorkbookFactory.create(inputStream);
		List<? extends PictureData> picList = workbook.getAllPictures();
		for(PictureData pic : picList) {
			ReadImageBean imageBean = new ReadImageBean();
			imageBean.setType(pic.getMimeType());
			imageBean.setSuffix(pic.suggestFileExtension());
			imageBean.setBytes(pic.getData());
			List.add(imageBean);
		}
		return List;
	}
	
	/**
	 * 创建文件
	 * 默认创建xlsx
	 * 默认数据写入第一个命名为first的sheet
	 * @param outputStream
	 * @param List
	 * @return
	 * @throws Exception
	 */
	public static void create(OutputStream outputStream, List<List<Object>> List) throws Exception {
		create(outputStream, List, "first");
	}
	
	/**
	 * 创建文件
	 * 默认创建xlsx
	 * @param outputStream
	 * @param List
	 * @param sheetName 指定sheet名称
	 * @return
	 * @throws Exception
	 */
	public static void create(OutputStream outputStream, List<List<Object>> List, String sheetName) throws Exception {
		Workbook workbook = new XSSFWorkbook();
		createSheet(workbook.createSheet(sheetName), List);
		workbook.write(outputStream);
		workbook.close();
	}
	
	/**
	 * 创建文件
	 * 默认创建xlsx
	 * @param outputStream
	 * @param Map key为sheet名称
	 * @return
	 * @throws Exception
	 */
	public static void create(OutputStream outputStream, Map<String, List<List<Object>>> map) throws Exception {
		Workbook workbook = new XSSFWorkbook();
		for(Entry<String, List<List<Object>>> entry : map.entrySet()) {
			createSheet(workbook.createSheet(entry.getKey()), entry.getValue());
		}
		workbook.write(outputStream);
		workbook.close();
	}
	
	/**
	 * 创建 sheet
	 * @param sheet
	 * @param List
	 * @return
	 * @throws Exception
	 */
	private static void createSheet(Sheet sheet, List<List<Object>> List) throws Exception {
		for (int i=0; i<List.size(); i++) {
			Row row = sheet.createRow(i);
			List<Object> list = List.get(i);
			for(int j=0; j<list.size(); j++) {
				writeCell(row.createCell(j), list.get(j));
			}
		}
	}
	
	/**
	 * 写入文件
	 * 向已有文件中固定位置写入
	 * 默认写入第一个sheet
	 * @param inputStream
	 * @param outputStream
	 * @param writeBeanList
	 * @return
	 * @throws Exception
	 */
	public static void write(InputStream inputStream, OutputStream outputStream, List<WriteBean> writeBeanList) throws Exception {
		write(inputStream, outputStream, writeBeanList, 0);
	}
	
	/**
	 * 写入文件
	 * 向已有文件中固定位置写入
	 * 按指定sheet下标写入
	 * @param inputStream
	 * @param outputStream
	 * @param writeBeanList
	 * @param sheetIndex
	 * @return
	 * @throws Exception
	 */
	public static void write(InputStream inputStream, OutputStream outputStream, List<WriteBean> writeBeanList, int sheetIndex) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		for(WriteBean writeBean : writeBeanList) {
			writeSheet(sheet, writeBean.getX(), writeBean.getY(), writeBean.getValue());
		}
		workbook.write(outputStream);
		workbook.close();
	}
	
	/**
	 * 写入文件
	 * 向已有文件中固定位置写入图片
	 * 默认下入第一个sheet
	 * @param inputStream
	 * @param outputStream
	 * @param writeImageBean
	 * @return
	 * @throws Exception
	 */
	public static void writeImage(InputStream inputStream, OutputStream outputStream, WriteImageBean writeImageBean) throws Exception {
		writeImage(inputStream, outputStream, writeImageBean, 0);
	}
	
	/**
	 * 写入文件
	 * 向已有文件中固定位置写入图片
	 * 按指定sheet下标写入
	 * @param inputStream
	 * @param outputStream
	 * @param writeImageBean
	 * @param sheetIndex
	 * @return
	 * @throws Exception
	 */
	public static void writeImage(InputStream inputStream, OutputStream outputStream, WriteImageBean writeImageBean, int sheetIndex) throws Exception {
		Workbook workbook = WorkbookFactory.create(inputStream);
		Sheet sheet = workbook.getSheetAt(sheetIndex);
		Drawing<?> drawing = sheet.createDrawingPatriarch();
		ClientAnchor anchor = (workbook instanceof XSSFWorkbook) ? 
			new XSSFClientAnchor(writeImageBean.getDx1(), writeImageBean.getDy1(), writeImageBean.getDx2(), writeImageBean.getDy2(), (short)writeImageBean.getCol1(), writeImageBean.getRow1(), (short)writeImageBean.getCol2(), writeImageBean.getRow2()) :
			new HSSFClientAnchor(writeImageBean.getDx1(), writeImageBean.getDy1(), writeImageBean.getDx2(), writeImageBean.getDy2(), (short)writeImageBean.getCol1(), writeImageBean.getRow1(), (short)writeImageBean.getCol2(), writeImageBean.getRow2());
		drawing.createPicture(anchor, workbook.addPicture(writeImageBean.getBytes(), Workbook.PICTURE_TYPE_JPEG)); // 此处图片类型先用固定值，亲测使用png图片可以写入成功
		workbook.write(outputStream);
		workbook.close();
	}

	
	/**
	 * 写入 sheet
	 * @param sheet
	 * @param rowIndex 行，从0开始
	 * @param cellIndex 列，从0开始
	 * @param value
	 * @return
	 */
	private static void writeSheet(Sheet sheet, int rowIndex, int cellIndex, Object value) {
		writeCell(sheet.getRow(rowIndex).getCell(cellIndex), value);
	}
	
	/**
	 * 写入 单元格
	 * @param cell
	 * @param value
	 * @return
	 */
	private static void writeCell(Cell cell, Object value) {
		if (Objects.isNull(value)) {
			cell.setCellValue(""); // null > ""
		}else if (value instanceof Integer) {
			cell.setCellValue((Integer)value);
		}else if (value instanceof Double) {
			cell.setCellValue((Double)value);
		}else if (value instanceof Boolean) {
			cell.setCellValue((Boolean)value);
		}else if (value instanceof Date) {
			cell.setCellValue((Date)value);
		}else if (value instanceof Calendar) {
			cell.setCellValue((Calendar)value);
		}else {
			cell.setCellValue(String.valueOf(value));
		}
	}
	
	
	
	/**
	 * 数据写入实体
	 * @author YangJie [2018年11月16日]
	 */
	public static class WriteBean{
		int x; // 横坐标
		int y; // 纵坐标
		Object value; // 内容
		public int getX() {
			return x;
		}
		public WriteBean setX(int x) {
			this.x = x;
			return this;
		}
		public int getY() {
			return y;
		}
		public WriteBean setY(int y) {
			this.y = y;
			return this;
		}
		public Object getValue() {
			return value;
		}
		public WriteBean setValue(Object value) {
			this.value = value;
			return this;
		}
	}	
	
	/**
	 * 写入图片实体
	 * @author YangJie [2018年11月16日]
	 */
	public static class WriteImageBean{
		private byte[] bytes; // 文件内容
		private int row1; // 左上角所在行
		private int col1; // 左上角所在列
		private int dx1; // 左上角横轴偏移量
		private int dy1; // 左上角纵轴偏移量
		private int row2; // 右下角所在行
		private int col2; // 右下角所在列
		private int dx2; // 右下角横轴偏移量
		private int dy2; // 右下角纵轴偏移量

		public byte[] getBytes() {
			return bytes;
		}
		public WriteImageBean setBytes(byte[] bytes) {
			this.bytes = bytes;
			return this;
		}
		public int getRow1() {
			return row1;
		}
		public WriteImageBean setRow1(int row1) {
			this.row1 = row1;
			return this;
		}
		public int getCol1() {
			return col1;
		}
		public WriteImageBean setCol1(int col1) {
			this.col1 = col1;
			return this;
		}
		public int getDx1() {
			return dx1;
		}
		public WriteImageBean setDx1(int dx1) {
			this.dx1 = dx1;
			return this;
		}
		public int getDy1() {
			return dy1;
		}
		public WriteImageBean setDy1(int dy1) {
			this.dy1 = dy1;
			return this;
		}
		public int getRow2() {
			return row2;
		}
		public WriteImageBean setRow2(int row2) {
			this.row2 = row2;
			return this;
		}
		public int getCol2() {
			return col2;
		}
		public WriteImageBean setCol2(int col2) {
			this.col2 = col2;
			return this;
		}
		public int getDx2() {
			return dx2;
		}
		public WriteImageBean setDx2(int dx2) {
			this.dx2 = dx2;
			return this;
		}
		public int getDy2() {
			return dy2;
		}
		public WriteImageBean setDy2(int dy2) {
			this.dy2 = dy2;
			return this;
		}		
	}
	
	/**
	 * 读取图片实体
	 * @author YangJie [2018年11月16日]
	 */
	public static class ReadImageBean{
		private byte[] bytes; // 文件内容
		private String type; // 文件类型
		private String suffix; // 文件后缀
		
		public byte[] getBytes() {
			return bytes;
		}
		public ReadImageBean setBytes(byte[] bytes) {
			this.bytes = bytes;
			return this;
		}
		public String getType() {
			return type;
		}
		public ReadImageBean setType(String type) {
			this.type = type;
			return this;
		}
		public String getSuffix() {
			return suffix;
		}
		public ReadImageBean setSuffix(String suffix) {
			this.suffix = suffix;
			return this;
		}
	}
		
}
