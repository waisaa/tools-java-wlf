package cn.wlftool.excel.entity;


import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;
import java.util.Map;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class ExcelWriteParam {

    // 目标文件路径
    private String filepath;

    // 模板文件路径
    private String tplFilepath;

    // excel文件中每个sheet(sheet名称或索引)对应的数据集合
    private Map<String, List<?>> sheetValues;

    // excel文件中每个sheet(sheet名称或索引)起始行索引（sheet对应的为空时按区域读取数据ExcelRegion）
    private Map<String, Integer> startRowIndexes;

    // 若有表头的话
    // excel文件中每个sheet(sheet名称或索引)对应的表头及起始行索引 <sheetName, [SheetHeader]>>>
    private Map<String, List<SheetHeader>> headers;

}
