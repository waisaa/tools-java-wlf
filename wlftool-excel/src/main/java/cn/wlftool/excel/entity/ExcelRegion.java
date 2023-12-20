package cn.wlftool.excel.entity;


import cn.wlftool.excel.constant.Const;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelRegion {

	// 起始行索引
	int startRowIndex() default 0;

	// 终止行索引(-1代表读取所有)
	int endRowIndex() default -1;

	// 起始列索引
	int startColIndex() default 0;

	// 终止列索引(-1代表读取所有)
	int endColIndex() default -1;

	// 是否加粗字体
	boolean bold() default false;

	// 是否为数值
	boolean numeric() default false;

	// 为数值时默认科学计数法表示
	boolean scientificNotation() default true;

	// 读写时该列的默认值
	String defaultValue() default Const.S_NULL;

	// 取值范围（校验）（除了所定义的类型除外的）
	String[] validValueRange() default {};

	// 错误提示（校验）
	String validErrorPrompt() default Const.S_NULL;

}
