package cn.itcast.common.excel;

import cn.itcast.common.excel.annotation.ExcelColumn;
import cn.itcast.common.excel.annotation.ExcelHeader;
import cn.itcast.common.excel.annotation.ExcelWarning;
import cn.itcast.common.excel.constants.ExcelType;
import cn.itcast.common.excel.model.CellColumnValue;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.collections.ListUtils;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.*;

/**
 * ClassName: ExcelUtils  
 * (Excel创建导入导出工具类)
 * @author zhangtian  
 * @version
 */
public class ExcelUtils {
	
	private ExcelUtils(){
		
	}
	
	/**
	 *
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param isBigData 是否开启大数据分页，true：是  false：否
	 * @param pageSize 分页每个页签显示的数据条数
	 * @return
	 * @author zhangtian
	 */
	public static final Workbook exportExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize) {
		return exportExcelDataData(null , appDatas, clazz, excelType, isBigData, pageSize) ;
	}

	/**
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 * 代码重构
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param isBigData 是否开启大数据分页，true：是  false：否
	 * @param pageSize 分页每个页签显示的数据条数
	 * @return
	 * @author zhangtian
	 */
	private static final Workbook exportExcelDataData(Workbook workbook, List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize) {
        Map<String, Object> results = excelDataResultMap(clazz) ;
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = excelLimitSheet(appDatas, isBigData, pageSize, null) ;

		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", isBigData) ;
		results.put("pageSize", pageSize) ;
		results.put("oldWorkbook", workbook) ;

		return createExcelData(results, excelType, clazz) ;
	}

    /**
     * 创建workbook,导出数据
     * @return
     */
	private static final Workbook createExcelData(Map<String, Object> results, ExcelType excelType, Class<?> clazz) {
        if(ExcelType.XLS.equals(excelType)) {
            return ExcelManager.createExcelManager().exportContainDataExcel_XLS(results, clazz) ;
        } else if(ExcelType.XLSX.equals(excelType)) {
            return ExcelManager.createExcelManager().exportContainDataExcel_XLSX(results, clazz);
        } else {
			/*
			 * 导出Excel应对一定量大数据策略2
			 * 分页签Sheet导出海量数据
			 * 导出数据后及时刷新内存
			 *
			 */
            return ExcelManager.createExcelManager().exportContainDataExcel_SXLSX(results, clazz) ;
        }
    }

	/**
	 * 创建workbook,导出数据
	 * @param results
	 * @param excelType
	 * @param clazzs
	 * @return
	 */
	private static final Workbook createVirtualRowExcelData(Map<String, Object> results, ExcelType excelType, Map<String, Class<?>> clazzs) {
		if(ExcelType.XLS.equals(excelType)) {
			return ExcelManager.createExcelManager().exportVirtualRollDataExcel_XLS(results, clazzs) ;
		} else if(ExcelType.XLSX.equals(excelType)) {
			return ExcelManager.createExcelManager().exportVirtualRollDataExcel_XLSX(results, clazzs) ;
		} else {
			/*
			 * 导出Excel应对一定量大数据策略2
			 * 分页签Sheet导出海量数据
			 * 导出数据后及时刷新内存
			 *
			 */
			return ExcelManager.createExcelManager().exportVirtualRollDataExcel_SXLSX(results, clazzs) ;
		}
	}

    /**
     * 拆分sheet
     * @return
     */
	private static final String[] excelLimitSheet(List<?> appDatas, boolean isBigData, int pageSize, String sheetNames) {
        // 大批量数据条件下的分割Sheet
        String[] sheetResult = null ;
        if(isBigData) {
            int size = appDatas.size() ;
            int sheetNums = size % pageSize == 0 ? size / pageSize : (size / pageSize +1) ;
            sheetResult = new String[sheetNums] ;
            if(sheetNums > 1) {
                for(int i = 0; i< sheetNums; i++) {
                    sheetResult[i] = ExcelManager.DEFAULT_SHEET_NAME + (i+1) ;
                }
            }
        } else {
            if(!"".equals(sheetNames) && null != sheetNames) {
                sheetResult = new String[] {sheetNames} ;
            }else{
                sheetResult = new String[]{ExcelManager.DEFAULT_SHEET_NAME+1} ;
            }
        }
	    return sheetResult ;
    }

	/**
	 * 向已有的workbook写数据
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param isBigData 是否开启大数据分页，true：是  false：否
	 * @param pageSize 分页每个页签显示的数据条数
	 * @return
	 * @author zhangtian
	 */
	private static final Workbook exportExcelDataOldWorkbook(Workbook workbook, List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize) {
		return exportExcelDataData(workbook , appDatas, clazz, excelType, isBigData, pageSize) ;
	}
	
	/**
	 * 
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param sheetNames 自定义Sheet页签的名称
	 * @return
	 * @author zhangtian
	 */
	public static final Workbook exportExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames) {
		return exportExcelDataData(null ,appDatas, clazz, excelType, sheetNames) ;
	}

	/**
	 * 数据导出，虚拟行，一键多值，导出一行对应多行的数据，数据模型为List<Map<Object,List<Object>>> datas = new ArrayList<Map<Object,List<Object>>>() ;
	 * @param datas
	 * @param clazzs
	 * @param excelType
	 * @param sheetNames
	 * @return
	 */
	public static final Workbook exportVirtualRowExcelData(List<Map<Object,List<Object>>> datas, ExcelType excelType, String sheetNames) {
		Map<String, Class<?>> clazzs = new HashMap<String, Class<?>>() ;
		// 根据spring核心包，获取参数泛型类型
		if(!CollectionUtils.isEmpty(datas)) {
			Map<Object, List<Object>> map0 = datas.get(0) ;
			if(MapUtils.isNotEmpty(map0)) {
				for(Map.Entry<Object, List<Object>> entry : map0.entrySet()) {
					Object key = entry.getKey() ;
					clazzs.put("keyClass", key.getClass()) ;
					List<Object> list = entry.getValue() ;
					if(CollectionUtils.isNotEmpty(list)) {
						Object value = list.get(0) ;
						clazzs.put("valueClass", value.getClass()) ;
					}else {
						clazzs.put("valueClass", null) ;
					}
				}
				return exportVirtualRowExcelDataData(null ,datas, clazzs, excelType, sheetNames) ;
			} else {
				return null ;
			}
		} else {
			return null ;
		}
	}

	/**
	 * 数据导出，虚拟行，一键多值
	 * @param workbook
	 * @param appDatas
	 * @param clazz
	 * @param excelType
	 * @param sheetNames
	 * @return
	 */
	private static final Workbook exportVirtualRowExcelDataData(Workbook workbook, List<Map<Object,List<Object>>> datas, Map<String, Class<?>> clazzs, ExcelType excelType, String sheetNames) {
		Map<String, Object> results = excelDataResultMap(clazzs.get("keyClass")) ;
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = excelLimitSheet(datas, false, 0, sheetNames) ;
		results.put("appDatas", datas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", false) ;
		results.put("pageSize", 0) ;
		results.put("oldWorkbook", workbook) ;

		return createVirtualRowExcelData(results, excelType, clazzs) ;
	}

    /**
     * 重构map参数，传递excel导出的数据参数
     * @return
     */
	private static final Map<String, Object> excelDataResultMap(Class<?> clazz) {
        Map<String, Object> results = new HashMap<String, Object>() ;
        Field[] fields = clazz.getDeclaredFields() ;

        // 保存标题
        if(clazz.isAnnotationPresent(ExcelHeader.class)) {
            ExcelHeader excelHeader = clazz.getAnnotation(ExcelHeader.class) ;
            results.put("headerName", excelHeader.headerName()) ;
        }

        // 保存警告信息
        if(clazz.isAnnotationPresent(ExcelWarning.class)) {
            ExcelWarning excelWarning = clazz.getAnnotation(ExcelWarning.class) ;
            results.put("warningInfo", excelWarning.warningInfo()) ;
        }

        // 保存列头信息
        List<CellColumnValue> list = new ArrayList<CellColumnValue>() ;
        for(Field field : fields) {
            if(field.isAnnotationPresent(ExcelColumn.class)) {
                CellColumnValue cellColumnValue = new CellColumnValue() ;
                ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class) ;
                cellColumnValue.setColumnKey(field.getName());
                if(excelColumn.columnName() == null || "".equals(excelColumn.columnName().trim())) {
                    cellColumnValue.setColumnValue(field.getName().toUpperCase());
                } else {
                    cellColumnValue.setColumnValue(excelColumn.columnName());
                }
                list.add(cellColumnValue) ;
            }
        }
        results.put("columnNames", list) ;

	    return results ;
    }

	/**
	 * 代码重构
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param sheetNames 自定义Sheet页签的名称
	 * @return
	 * @author zhangtian
	 */
	private static final Workbook exportExcelDataData(Workbook workbook ,List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames) {
        Map<String, Object> results = excelDataResultMap(clazz) ;
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = excelLimitSheet(appDatas, false, 0, sheetNames) ;
		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", false) ;
		results.put("pageSize", 0) ;
		results.put("oldWorkbook", workbook) ;

        return createExcelData(results, excelType, clazz) ;
	}

	/**
	 * 向已有的workbook写数据
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param sheetNames 自定义Sheet页签的名称
	 * @return
	 * @author zhangtian
	 */
	private static final Workbook exportExcelDataOldWorkbook(Workbook workbook ,List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames) {
		return exportExcelDataData(workbook ,appDatas, clazz, excelType, sheetNames) ;
	}

	/**
	 *
	 * importExcelData: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 * 
	 * @param filePath
	 * @param excelType Excel类型
	 * @param clazz 携带注解的Bean字节码文件
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public static final List<Object> importAllExcelData(String filePath, ExcelType excelType, Class<?> clazz) throws FileNotFoundException, IOException, SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		List<Object> list = new ArrayList<Object>() ;
		
		Iterator<Sheet> it = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filePath));
			it = excelManager.getHSSFWorkbook(fs).sheetIterator();
		} else {
			it = excelManager.getXSSFWorkbook(filePath).sheetIterator() ;
		}
		
		while(it.hasNext()) {
			list.addAll(excelManager.importExcelData(clazz, it.next())) ;
		}
		return list ;
	}

	/**
	 *
	 * importExcelData: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 * 
	 * @param filePath
	 * @param excelType Excel类型
	 * @param clazz 携带注解的Bean字节码文件
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public static final List<Object> importAllExcelData(InputStream in, ExcelType excelType, Class<?> clazz) throws FileNotFoundException, IOException, SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		List<Object> list = new ArrayList<Object>() ;
		
		Iterator<Sheet> it = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			it = excelManager.getHSSFWorkbook(fs).sheetIterator();
		} else {
			it = excelManager.getXSSFWorkbook(in).sheetIterator() ;
		}
		
		while(it.hasNext()) {
			list.addAll(excelManager.importExcelData(clazz, it.next())) ;
		}
		return list ;
	}

	/**
	 * 
	 *  importExcelData:(根据Sheet名称指定遍历). 
	 *  @return_type:List<Object>
	 *  @author zhangtian  
	 *  @param in
	 *  @param clazz
	 *  @param sheetNames
	 *  @return
	 *  @throws IOException
	 *  @throws SecurityException
	 *  @throws NoSuchFieldException
	 *  @throws InstantiationException
	 *  @throws IllegalAccessException
	 */
	public static final List<Object> importExcelData(InputStream in, ExcelType excelType, Class<?> clazz, String... sheetNames) throws IOException,
		SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
	
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		if(ArrayUtils.isEmpty(sheetNames)) {
			throw new RuntimeException("请指定页签Sheet名称...") ;
		}
		
		Workbook workbook = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			workbook = excelManager.getHSSFWorkbook(fs);
		} else {
			workbook = excelManager.getXSSFWorkbook(in);
		}
		
		for(String name : sheetNames) {
			if(workbook.getSheet(name) == null) {
				throw new RuntimeException("Sheet页签指定的名称不存在......") ;
			}
			continue ;
		}
		
		Sheet[] sheets = new Sheet[sheetNames.length] ;
		int i = 0; 
		for(String sheetName : sheetNames) {
			sheets[i] = workbook.getSheet(sheetName) ;
			i++ ;
		}
		
		return excelManager.importExcelData(clazz, sheets);
	}
	
	/**
	 * 
	 * importExcelDate: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 *  importExcelData:(根据Sheet索引指定遍历). 
	 *  @return_type:List<Object>
	 *  @author zhangtian  
	 *  @param in
	 *  @param clazz
	 *  @param sheetNames
	 *  @return
	 *  @throws IOException
	 *  @throws SecurityException
	 *  @throws NoSuchFieldException
	 *  @throws InstantiationException
	 *  @throws IllegalAccessException
	 */
	public static final List<Object> importExcelData(InputStream in, ExcelType excelType, Class<?> clazz, int... sheetIndexes) throws IOException,
		SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		
		if(ArrayUtils.isEmpty(sheetIndexes)) {
			throw new RuntimeException("请指定页签Sheet索引...") ;
		}
		
		Workbook workbook = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			workbook = excelManager.getHSSFWorkbook(fs);
		} else {
			workbook = excelManager.getXSSFWorkbook(in);
		} 
		
		int sheetNum = workbook.getNumberOfSheets() ;
		for(int index : sheetIndexes) {
			if((index+1) > sheetNum) {
				throw new RuntimeException("Sheet页签下标越界......") ;
			}
			continue ;
		}
		
		Sheet[] sheets = new Sheet[sheetIndexes.length] ;
		int i = 0; 
		for(int sheetIndex : sheetIndexes) {
			sheets[i] = workbook.getSheetAt(sheetIndex) ;
			i++ ;
		}
		
		return excelManager.importExcelData(clazz, sheets);
	}
	
	// ================================ Excel错误模板导出工具
	/**
	 * 
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param isBigData 是否开启大数据分页，true：是  false：否
	 * @param pageSize 分页每个页签显示的数据条数
	 * @return
	 * @author zhangtian
	 */
	public static final Workbook exportErrorExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize) {
        Map<String, Object> results = excelDataResultMap(clazz) ;
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = excelLimitSheet(appDatas, isBigData, pageSize, null);
		
		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", isBigData) ;
		results.put("pageSize", pageSize) ;

		return createExcelErrorData(results, excelType, clazz) ;
	}

    /**
     * 创建workbook,导出错误数据
     * @return
     */
    private static final Workbook createExcelErrorData(Map<String, Object> results, ExcelType excelType, Class<?> clazz) {
        if(ExcelType.XLS.equals(excelType)) {
            return ExcelManager.createExcelManager().exportContainErrorDataExcel_XLS(results, clazz) ;
        } else if(ExcelType.XLSX.equals(excelType)) {
            return ExcelManager.createExcelManager().exportContainErrorDataExcel_XLSX(results, clazz);
        } else {
			/*
			 * 导出Excel应对一定量大数据策略2
			 * 分页签Sheet导出海量数据
			 * 导出数据后及时刷新内存
			 *
			 */
            return ExcelManager.createExcelManager().exportContainErrorDataExcel_SXLSX(results, clazz) ;
        }
    }
	
	/**
	 * 
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param sheetNames 自定义Sheet页签的名称
	 * @return
	 * @author zhangtian
	 */
	public static final Workbook exportErrorExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames) {
		Map<String, Object> results = excelDataResultMap(clazz) ;
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = excelLimitSheet(appDatas, false, 0, sheetNames);

		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", false) ;
		results.put("pageSize", 0) ;

        return createExcelErrorData(results, excelType, clazz) ;
	}

	/**
	 * 获取workbook对象
	 * @param excelType
	 * @param in
	 * @return
	 * @throws IOException
	 */
	public static final Workbook getWorkbook(ExcelType excelType, InputStream in) throws IOException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		Workbook workbook = null ;
        try {
			if(ExcelType.XLS.equals(excelType)) {
				if(in != null) {
					POIFSFileSystem poifsFileSystem = null ;
					try {
						poifsFileSystem = new POIFSFileSystem(in) ;
						workbook = excelManager.getHSSFWorkbook(poifsFileSystem) ;
					} finally {
						if(poifsFileSystem != null) {
							poifsFileSystem.close();
						}
					}
				}else{
					workbook = excelManager.getHSSFWorkbook() ;
				}
			}else if(ExcelType.XLSX.equals(excelType)) {
				if(in != null) {
					workbook = excelManager.getXSSFWorkbook(in) ;
				}else {
					workbook = excelManager.getXSSFWorkbook() ;
				}
			}else {
				if(in != null) {
					workbook = excelManager.getSXSSFWorkbook(in) ;
				}else {
					workbook = excelManager.getSXSSFWorkbook() ;
				}
			}
		}finally {
			if(in != null) {
				in.close();
			}
		}

		return workbook ;
	}

	/**
	 * 扩展导出功能 增强在原有Excel基础上重写
	 * @param oldWorkbook
	 * @param appDatas
	 * @param clazz
	 * @param excelType
	 * @param isBigData
	 * @param pageSize
	 * @return
	 */
	public static final Workbook exportExcelDataToOldWorkbook(Workbook oldWorkbook, List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize){
		if(oldWorkbook != null) {
			return exportExcelDataOldWorkbook(oldWorkbook,appDatas, clazz, excelType, isBigData, pageSize) ;
		}else {
			return exportExcelData(appDatas, clazz, excelType, isBigData, pageSize) ;
		}
	}

	/**
	 * 扩展导出功能 增强在原有Excel基础上重写
	 * @param oldWorkbook
	 * @param appDatas
	 * @param clazz
	 * @param excelType
	 * @param isBigData
	 * @param pageSize
	 * @return
	 */
	public static final Workbook exportExcelDataToOldWorkbook(Workbook oldWorkbook, List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames){
		if(oldWorkbook != null) {
			return exportExcelDataOldWorkbook(oldWorkbook, appDatas, clazz, excelType, sheetNames) ;
		}else {
			return exportExcelData(appDatas, clazz, excelType, sheetNames) ;
		}
	}
}
