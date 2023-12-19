package cn.wlftool.excel.entity;


import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelCol {

	// 读时列索引(-1代表该列设置为默认值)
	int readIndex() default -1;

	// 写时列索引(-1代表该列设置为默认值)
	int writeStartIndex() default -1;

	// 读写时该列的默认值
	String defaultValue() default Const.S_NULL;

	// 该列是否为数值
	boolean numeric() default false;

	// 该列为数值时默认科学计数法表示
	boolean scientificNotation() default true;

	// 取值范围（校验）
	String[] validValueRange() default {};

	// 错误提示（校验）
	String validErrorPrompt() default Const.S_NULL;

}
