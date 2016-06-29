package cn.itcast.common.excel.annotation;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * ClassName: ExcelWarning  
 * (Excel注意事项列表)
 * @author zhangtian  
 * @version
 */
//注解范围   类上注解
@Target({ElementType.TYPE})
//注解加载时机  运行时加载
@Retention(RetentionPolicy.RUNTIME)
//是否生成注解文档
@Documented
public @interface ExcelWarning {
	String[] warningInfo() default "" ;
}
