package cn.wlftool.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.List;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class SheetHeader {

    // 表头所在行索引
    private int rowIndex;

    // 起始列索引
    private int startColIndex;

    // 表头集合
    private List<String> headers;

}
