package cn.itcast.common.excel;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.List;
import java.util.Map;

import org.apache.commons.codec.binary.StringUtils;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.itcast.common.excel.annotation.ExcelColumn;
import cn.itcast.common.excel.constants.CommentType;
import cn.itcast.common.excel.model.CellColumnValue;

/**
 * 
 * ClassName:ExcelManager Function: 操作Excel工具类 功能 : Excel文件的导入 Excel文件的导出
 * 
 * @author zhangtian
 * @Date 2014 2014年8月8日 下午7:15:51
 *
 */
public class ExcelManager {
	public static final String DEFAULT_SHEET_NAME = "sheet";
	// === Workbook对象
	private Workbook workbook = null;
	// === Excel文件的头部标题样式
	private CellStyle headerStyle = null;
	// === Excel文件的第二行提示信息样式
	private CellStyle warnerStyle = null;
	// === Excel文件列头的样式
	private CellStyle titleStyle = null;
	// === Excel文件的数据样式
	private CellStyle dataStyle = null;
	// === Excel文件中的错误数据的显示样式
	private CellStyle errorDataStyle = null;
	
	private ExcelManager() {
		
	}
	
	// ======================================== Excel 公共方法调用
	// =============================================
	/*
	 * 获取HSSFWorkbook对象
	 */
	private HSSFWorkbook getHSSFWorkbook() {
		return new HSSFWorkbook();
	}
	
	private HSSFWorkbook getHSSFWorkbook(POIFSFileSystem in) throws IOException {
		return new HSSFWorkbook(in);
	}
	
	/*
	 * 获取XSSFWorkbook对象
	 */
	private XSSFWorkbook getXSSFWorkbook() {
		return new XSSFWorkbook();
	}
	
	private XSSFWorkbook getXSSFWorkbook(InputStream in) throws IOException {
		return new XSSFWorkbook(in);
	}
	
	/*
	 * 获取Excel管理器
	 */
	public static final ExcelManager createExcelManager() {
		return ExcelManegerTool.INSTANCE ;
	}
	
	private static class ExcelManegerTool {
		private final static ExcelManager INSTANCE = new ExcelManager() ;
	}
	
	// ======================================= 创建公共样式
	// ==============================================
	/*
	 * 设置Excel文件的头部标题的样式
	 */
	private void setHeaderCellStyles(Workbook workbook) {
		headerStyle = workbook.createCellStyle();

		// === 设置边框
		headerStyle.setBorderBottom(CellStyle.BORDER_THIN);
		headerStyle.setBorderLeft(CellStyle.BORDER_THIN);
		headerStyle.setBorderRight(CellStyle.BORDER_THIN);
		headerStyle.setBorderTop(CellStyle.BORDER_THIN);

		// === 设置背景色
		headerStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		headerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// === 设置居中
		headerStyle.setAlignment(CellStyle.ALIGN_CENTER);
		headerStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		// === 设置字体
		Font font = workbook.createFont();
		font.setFontName("粗体");

		// === 设置字体大小
		font.setFontHeightInPoints((short) 16);

		// === 设置粗体显示
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// === 选择需要用到的字体格式
		headerStyle.setFont(font);
		// === 设置自动换行
		// headerStyle.setWrapText(true) ;
		// sheet.autoSizeColumn((short)0, true); // === 调整第一列宽度
	}
	
	/*
	 * 设置Excel文件的第二列的注意事项提示信息的样式
	 */
	private void setWarnerCellStyles(Workbook workbook) {
		warnerStyle = workbook.createCellStyle();

		// === 设置边框
		warnerStyle.setBorderBottom(CellStyle.BORDER_THIN);
		warnerStyle.setBorderLeft(CellStyle.BORDER_THIN);
		warnerStyle.setBorderRight(CellStyle.BORDER_THIN);
		warnerStyle.setBorderTop(CellStyle.BORDER_THIN);

		// === 设置背景色
		warnerStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		warnerStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// === 设置左对齐
		warnerStyle.setAlignment(CellStyle.ALIGN_LEFT);

		// === 设置字体
		Font font = workbook.createFont();
		font.setFontName("宋体");

		// === 设置字体大小
		font.setFontHeightInPoints((short) 10);

		// === 设置粗体显示
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// === 设置字体颜色
		font.setColor(HSSFColor.RED.index);

		// === 选择需要用到的字体格式
		warnerStyle.setFont(font);

		// === 设置自动换行
		warnerStyle.setWrapText(true);
	}
	
	/*
	 * 设置Excel文件的列头样式
	 */
	private void setTitleCellStyles(Workbook workbook) {
		titleStyle = workbook.createCellStyle();
		// === 设置边框
		titleStyle.setBorderBottom(CellStyle.BORDER_THIN);
		titleStyle.setBorderLeft(CellStyle.BORDER_THIN);
		titleStyle.setBorderRight(CellStyle.BORDER_THIN);
		titleStyle.setBorderTop(CellStyle.BORDER_THIN);

		// === 设置背景色
		titleStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		titleStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// === 设置居中
		titleStyle.setAlignment(CellStyle.ALIGN_CENTER);
		titleStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

		// === 设置字体
		Font font = workbook.createFont();
		font.setFontName("粗体");

		// === 设置字体大小
		font.setFontHeightInPoints((short) 12);

		// === 设置粗体显示
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);

		// === 选择需要用到的字体格式
		titleStyle.setFont(font);

		// === 设置自动换行
		// titleStyle.setWrapText(true) ;
	}
	
	/*
	 * 设置Excel文件的数据样式
	 */
	private void setDataCellStyles(Workbook workbook) {
		dataStyle = workbook.createCellStyle();
		// === 设置单元格格式为文本格式
		DataFormat dataFormat = workbook.createDataFormat();
		dataStyle.setDataFormat(dataFormat.getFormat("@"));
		dataStyle.setBorderBottom(CellStyle.BORDER_THIN);
		dataStyle.setBorderLeft(CellStyle.BORDER_THIN);
		dataStyle.setBorderRight(CellStyle.BORDER_THIN);
		dataStyle.setBorderTop(CellStyle.BORDER_THIN);

		// === 设置背景色
		dataStyle.setFillForegroundColor(HSSFColor.LIGHT_GREEN.index);
		dataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// === 设置居中
		dataStyle.setAlignment(CellStyle.ALIGN_LEFT);

		// === 设置字体
		Font font = workbook.createFont();
		font.setFontName("宋体");

		// === 设置字体大小
		font.setFontHeightInPoints((short) 11);

		// === 选择需要用到的字体格式
		dataStyle.setFont(font);

		// === 设置自动换行
		// dataStyle.setWrapText(true) ;
	}
	
	/*
	 * 错误数据重新导入Excel中的样式
	 */
	private void setErrorDataStyle(Workbook workbook) {
		errorDataStyle = workbook.createCellStyle();
		// === 设置边框颜色
		errorDataStyle.setBottomBorderColor(HSSFColor.RED.index);
		errorDataStyle.setLeftBorderColor(HSSFColor.RED.index);
		errorDataStyle.setRightBorderColor(HSSFColor.RED.index);
		errorDataStyle.setTopBorderColor(HSSFColor.RED.index);

		// === 设置边框
		errorDataStyle.setBorderBottom(CellStyle.BORDER_THIN);
		errorDataStyle.setBorderLeft(CellStyle.BORDER_THIN);
		errorDataStyle.setBorderRight(CellStyle.BORDER_THIN);
		errorDataStyle.setBorderTop(CellStyle.BORDER_THIN);

		// === 设置背景色
		errorDataStyle.setFillForegroundColor(HSSFColor.ROSE.index);
		errorDataStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);

		// === 设置居中
		errorDataStyle.setAlignment(CellStyle.ALIGN_LEFT);

		// === 设置字体
		Font font = workbook.createFont();
		font.setFontName("宋体");

		// === 设置字体大小
		font.setFontHeightInPoints((short) 11);

		// === 选择需要用到的字体格式
		errorDataStyle.setFont(font);
	}
	
	// ======================================= 创建批注对象
	// ==============================================
	// 批注
	private Drawing getDrawing (Sheet sheet) {
		Drawing drawing = sheet.createDrawingPatriarch();
		return drawing ;
	}
	
	// ======================================= 创建公共数据
	// ==============================================
	// 创建单元格表头数据
	private void createAppRowHeaderData(String headerTitle, int startFlag, Integer cellHeaderNum, Sheet... sheets) {
		// 如果表头的内容为空，则判断该单元格不需要表头，则直接跳过
		if(headerTitle == null || "".equals(headerTitle.trim())) 
			return ;
		
		// 没有列头，单元格没有创建的意义
		if(cellHeaderNum == 0)
			return ;
		
		if(sheets != null) {
			if(sheets.length > 0) {
				// 循环创建页签Sheet的标题，考虑到大数据
				for(Sheet sheet : sheets) {
					// 创建表头数据,创建第一行,起始行标记
					Row row = sheet.createRow(startFlag);
					// 设置行高
					row.setHeight((short) 800);

					// 创建第一列单元格并设置样式
					Cell headerCell = row.createCell(0);
					
					// 设置批注
					Drawing drawing = getDrawing(sheet) ;
					Comment comment = null ;
					// === 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
					if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
						comment = drawing.createCellComment(new HSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) (cellHeaderNum - 1), (short)3, 3, (short)5, 6)) ;
						comment.setString(new HSSFRichTextString(CommentType.EXCEL_HEADER.name()));
					} else {
						comment = drawing.createCellComment(new XSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) (cellHeaderNum - 1), (short)3, 3, (short)5, 6)) ;
						comment.setString(new XSSFRichTextString(CommentType.EXCEL_HEADER.name()));
					}
					
					// 输入批注信息
					comment.setAuthor("zhangtian@fengyuntec.com");
				    //将批注添加到单元格对象中
					headerCell.setCellComment(comment);
					if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
						headerCell.setCellValue(new HSSFRichTextString(headerTitle));
					} else {
						headerCell.setCellValue(new XSSFRichTextString(headerTitle));
					}
					headerCell.setCellStyle(headerStyle);

					// 循环创建空的单元格，合并单元格需要
					for (int i = 1; i < cellHeaderNum; i++) {
						headerCell = row.createCell(i);
						headerCell.setCellStyle(headerStyle);
					}
					// === 合并头部单元格 参数：firstRow, lastRow, firstCol, lastCol
					sheet.addMergedRegion(new CellRangeAddress((short) startFlag, (short) startFlag, (short) 0, (short) (cellHeaderNum - 1)));
					// === 设置单元格自动列宽，中文支持较好
					// sheet.setColumnWidth(0, headerTitle.getBytes().length*2*256);
					// 设置自动列宽
					for (int i = 0; i < cellHeaderNum; i++) {
						sheet.autoSizeColumn((short) i, true);
					}
				}
			}
		}
	}
	
	// 创建警告头信息
	private void createAppWaringData(String[] warningInfo, int startFlag, Integer cellHeaderNum, Sheet... sheets) {
		// 如果表头的内容为空，则判断该单元格不需要表头，则直接跳过
		if(warningInfo == null || warningInfo.length == 0) 
			return ;
		
		// 没有列头，单元格没有创建的意义
		if(cellHeaderNum == 0)
			return ;
		
		if(sheets != null) {
			if(sheets.length > 0) {
				// 循环创建页签Sheet的标题，考虑到大数据
				for(Sheet sheet : sheets) {
					// 创建表头数据,创建第一行,起始行标记
					Row row = sheet.createRow(startFlag);
					// 设置行高
					row.setHeight((short) 1800);

					// 创建第一列单元格并设置样式
					Cell warningCell = row.createCell(0);
					
					// 设置批注
					Drawing drawing = getDrawing(sheet) ;
					Comment comment = null ;
					// === 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
					if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
						comment = drawing.createCellComment(new HSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) (cellHeaderNum - 1), (short)3, 3, (short)5, 6)) ;
						comment.setString(new HSSFRichTextString(CommentType.EXCEL_WARING.name()));
					} else {
						comment = drawing.createCellComment(new XSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) (cellHeaderNum - 1), (short)3, 3, (short)5, 6)) ;
						comment.setString(new XSSFRichTextString(CommentType.EXCEL_WARING.name()));
					}
					
					// 输入批注信息
					comment.setAuthor("zhangtian@fengyuntec.com");
				    //将批注添加到单元格对象中
					warningCell.setCellComment(comment);
					
					String warnResult = "" ;
					for(String warning : warningInfo) {
						warnResult += warning + "\r\n" ;
					}
					
					if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
						warningCell.setCellValue(new HSSFRichTextString(warnResult));
					} else {
						warningCell.setCellValue(new XSSFRichTextString(warnResult));
					}
					warningCell.setCellStyle(warnerStyle);

					// 循环创建空的单元格，合并单元格需要
					for (int i = 1; i < cellHeaderNum; i++) {
						warningCell = row.createCell(i);
						warningCell.setCellStyle(warnerStyle);
					}
					// === 合并头部单元格 参数：firstRow, lastRow, firstCol, lastCol
					sheet.addMergedRegion(new CellRangeAddress((short) startFlag, (short) startFlag, (short) 0, (short) (cellHeaderNum - 1)));
					// === 设置单元格自动列宽，中文支持较好
					// sheet.setColumnWidth(0, headerTitle.getBytes().length*2*256);
					// 设置自动列宽
					for (int i = 0; i < cellHeaderNum; i++) {
						sheet.autoSizeColumn((short) i, true);
					}
				}
			}
		}
	}
	
	/*
	 * 创建列头
	 */
	private void createAppRowCellHeaderData(int startFlag, List<CellColumnValue> cellHeader, Class<?> clazz, Sheet... sheets) {
		// 循环创建页签Sheet的列头，考虑到大数据
		if(sheets != null) {
			if(sheets.length > 0) {
				for(Sheet sheet : sheets) {
					// 创建列头行
					Row row = sheet.createRow(startFlag);
					row.setHeight((short) 500);

					Cell cellHeaderCell = null;
					if (cellHeader != null) {
						if(!cellHeader.isEmpty()) {
							Drawing drawing = getDrawing(sheet) ;
							// 设置批注
							Comment comment = null ;
							// 循环创建列头
							for (int i = 0; i < cellHeader.size(); i++) {
								cellHeaderCell = row.createCell(i);
								if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
									cellHeaderCell.setCellValue(new HSSFRichTextString(cellHeader.get(i).getColumnValue()));
								} else {
									cellHeaderCell.setCellValue(new XSSFRichTextString(cellHeader.get(i).getColumnValue()));
								}
								cellHeaderCell.setCellStyle(titleStyle);
								// === 设置列宽
								sheet.setColumnWidth(i, (short) 7000);
								
								// === 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
								if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
									comment = drawing.createCellComment(new HSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) i, (short)3, 3, (short)5, 6)) ;
									comment.setString(new HSSFRichTextString(cellHeader.get(i).getColumnKey()));
								} else {
									comment = drawing.createCellComment(new XSSFClientAnchor((short) startFlag, (short) startFlag, (short) startFlag, (short) i, (short)3, 3, (short)5, 6)) ;
									comment.setString(new XSSFRichTextString(cellHeader.get(i).getColumnKey()));
								}
								
								// 输入批注信息
								comment.setAuthor("zhangtian@fengyuntec.com");
							    //将批注添加到单元格对象中
								cellHeaderCell.setCellComment(comment);
							}
						}
					}
				}
			}
		}
	}
	
	/*
	 * 创建数据
	 */
	private void createAppRowHasData(int startFlag, List<Object> appData, Class<?> clazz, Integer cellHeaderNum,boolean isBigData, int pageSize, Sheet... sheets) {

		Row row = null;
		Cell cellAppDataCell = null;
		ExcelColumn excelColumn = null;
		if (cellHeaderNum != 0) {
			if(CollectionUtils.isNotEmpty(appData)) {
				if(isBigData) {
					int totalSize = appData.size() ;
					int start = 0 ;
					pageSize = pageSize >= totalSize ? totalSize : pageSize ;
					int end = pageSize ;
					for(Sheet sheet : sheets) {
						System.out.println(start+"===>"+end);
						// === 行记录数
						int k = 0 ;
						for (int i = start; i < end; i++) {
							// === 列记录数
							row = sheet.createRow(k+startFlag);
							k++;
							Object o = appData.get(i);
							Field[] fields = o.getClass().getDeclaredFields();
							int j = 0;
							for (Field field : fields) {
								if (field.isAnnotationPresent(ExcelColumn.class)) {
									field.setAccessible(true);
									excelColumn = field.getAnnotation(ExcelColumn.class);
									try {
										cellAppDataCell = row.createCell(j);
										if(StringUtils.equals(excelColumn.autoIncrement(), "Y")){
											if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
												cellAppDataCell.setCellValue(new HSSFRichTextString((k+1)+""));
											} else {
												cellAppDataCell.setCellValue(new XSSFRichTextString((k+1)+""));
											}
										}else{
											Object value = field.get(o);
											if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
												if (value != null) {
													cellAppDataCell.setCellValue(new HSSFRichTextString(
															cn.itcast.common.excel.utils.StringUtils.replaceEscapeChar(value.toString())));
												} else {
													cellAppDataCell.setCellValue(new HSSFRichTextString(""));
												}
											} else {
												if (value != null) {
													cellAppDataCell.setCellValue(new XSSFRichTextString(
															cn.itcast.common.excel.utils.StringUtils.replaceEscapeChar(value.toString())));
												} else {
													cellAppDataCell.setCellValue(new XSSFRichTextString(""));
												}
											}
											
										}
										cellAppDataCell.setCellStyle(dataStyle);
										j++;
									} catch (Exception e) {
										e.printStackTrace();
									}
								}
							}
						}
						
						start = start + pageSize ;
						end = end + pageSize ;
						if(end >= totalSize) {
							end = totalSize ;
						}
					}
				} else {
					for(Sheet sheet : sheets) {
						// === 行记录数
						for (int i = 0; i < appData.size(); i++) {
							// === 列记录数
							row = sheet.createRow(i + startFlag);
							Object o = appData.get(i);
							Field[] fields = o.getClass().getDeclaredFields();
							int j = 0;
							for (Field field : fields) {
								if (field.isAnnotationPresent(ExcelColumn.class)) {
									field.setAccessible(true);
									excelColumn = field.getAnnotation(ExcelColumn.class);
									try {
										cellAppDataCell = row.createCell(j);
										if(StringUtils.equals(excelColumn.autoIncrement(), "Y")){
											if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
												cellAppDataCell.setCellValue(new HSSFRichTextString((i+1)+""));
											} else {
												cellAppDataCell.setCellValue(new XSSFRichTextString((i+1)+""));
											}
										}else{
											Object value = field.get(o);
											if(sheet.getWorkbook().getClass().isAssignableFrom(HSSFWorkbook.class)) {
												if (value != null) {
													cellAppDataCell.setCellValue(new HSSFRichTextString(
															cn.itcast.common.excel.utils.StringUtils.replaceEscapeChar(value.toString())));
												} else {
													cellAppDataCell.setCellValue(new HSSFRichTextString(""));
												}
											} else {
												if (value != null) {
													cellAppDataCell.setCellValue(new XSSFRichTextString(
															cn.itcast.common.excel.utils.StringUtils.replaceEscapeChar(value.toString())));
												} else {
													cellAppDataCell.setCellValue(new XSSFRichTextString(""));
												}
											}
										}
										cellAppDataCell.setCellStyle(dataStyle);
										j++;
									} catch (Exception e) {
										e.printStackTrace();
									}
								}
							}
						}
					}
				}
			}
		}
	}
	
	// =========================================== 创建数据导入导出
	// ===========================================
	// === 导出Excel的表格
	@SuppressWarnings({ "unchecked"})
	public Workbook exportContainDataExcel_XLS(Map<String, Object> results, Class<?> clazz) {
		// ======================== 页签创建 ==========================
		// === 获取HSSFWorkbook对象
		workbook = getHSSFWorkbook();

		String[] sheetNames = (String[]) results.get("sheetNames") ;
		Sheet[] sheets = new Sheet[sheetNames.length] ;
		for(int i = 0; i<sheetNames.length; i++) {
			sheets[i] = workbook.createSheet(sheetNames[i]);
		}
		// ========================= 样式设置 =========================
		// === 设置表头样式
		setHeaderCellStyles(workbook);
		// 设置警告信息样式
		setWarnerCellStyles(workbook);
		// === 设置列头样式
		setTitleCellStyles(workbook);
		// === 设置数据样式
		setDataCellStyles(workbook);

		// ========================= 数据创建 ==========================
		// === 创建标题数据
		int startFlag = 0 ;
		if(results.get("headerName") != null) {
			createAppRowHeaderData(results.get("headerName").toString(),startFlag,((List<String>) results.get("columnNames")).size(), sheets);
			startFlag++ ;
		}
		
		// 创建警告头信息
		if(results.get("warningInfo") != null) {
			createAppWaringData((String[])results.get("warningInfo"),startFlag, ((List<String>) results.get("columnNames")).size(),sheets) ;
			startFlag++ ;
		}
		
		// === 创建列头数据信息
		if(results.get("columnNames") != null) {
			createAppRowCellHeaderData(startFlag, (List<CellColumnValue>) results.get("columnNames"), clazz, sheets);
			startFlag++ ;
		}
		// === 为空模板创建初始化数据 空数据样式
		createAppRowHasData(startFlag, (List<Object>) results.get("appDatas"), clazz,
				((List<String>) results.get("columnNames")).size(),(boolean)results.get("isBigData"), 
				(int)results.get("pageSize"),sheets);
		return workbook;
		// ========================= 文件输出 ==========================
		// FileOutputStream out = new FileOutputStream(filePath);
		// workbook.write(out);
		// out.close();
	}
	
	// === 导出Excel的表格
	@SuppressWarnings({ "unchecked"})
	public Workbook exportContainDataExcel_XLSX(Map<String, Object> results, Class<?> clazz) {
		// ======================== 页签创建 ==========================
		// === 获取HSSFWorkbook对象
		workbook = getXSSFWorkbook();

		String[] sheetNames = (String[]) results.get("sheetNames") ;
		Sheet[] sheets = new Sheet[sheetNames.length] ;
		for(int i = 0; i<sheetNames.length; i++) {
			sheets[i] = workbook.createSheet(sheetNames[i]);
		}
		// ========================= 样式设置 =========================
		// === 设置表头样式
		setHeaderCellStyles(workbook);
		// 设置警告信息样式
		setWarnerCellStyles(workbook);
		// === 设置列头样式
		setTitleCellStyles(workbook);
		// === 设置数据样式
		setDataCellStyles(workbook);

		// ========================= 数据创建 ==========================
		// === 创建标题数据
		int startFlag = 0 ;
		if(results.get("headerName") != null) {
			createAppRowHeaderData(results.get("headerName").toString(),startFlag,((List<String>) results.get("columnNames")).size(), sheets);
			startFlag++ ;
		}
		
		// 创建警告头信息
		if(results.get("warningInfo") != null) {
			createAppWaringData((String[])results.get("warningInfo"),startFlag, ((List<String>) results.get("columnNames")).size(),sheets) ;
			startFlag++ ;
		}
		
		// === 创建列头数据信息
		if(results.get("columnNames") != null) {
			createAppRowCellHeaderData(startFlag, (List<CellColumnValue>) results.get("columnNames"), clazz, sheets);
			startFlag++ ;
		}
		// === 为空模板创建初始化数据 空数据样式
		createAppRowHasData(startFlag, (List<Object>) results.get("appDatas"), clazz,
				((List<String>) results.get("columnNames")).size(),(boolean)results.get("isBigData"), 
				(int)results.get("pageSize"),sheets);
		return workbook;
		// ========================= 文件输出 ==========================
		// FileOutputStream out = new FileOutputStream(filePath);
		// workbook.write(out);
		// out.close();
	}
	
}
