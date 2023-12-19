package cn.wlftool.excel.entity;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Map;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelReadParam {

    // 目标文件路径
    private String filepath;

    // excel文件中每个sheet(sheet名称或索引)对应的数据对象
    private Map<String, Class<?>> sheetObject;

    // excel文件中每个sheet(sheet名称或索引，同sheetObject一致)对应的起始行索引（sheet对应的为空时按区域读取数据ExcelRegion）
    private Map<String, Integer> startRowIndexes;

    // 校验出错时，错误提示文件的生成路径
    private String validErrorPromptFilepath;

}
