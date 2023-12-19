package cn.wlftool.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class SheetRegion {

    // 起止行索引
    private int startRowIndex;
    private int endRowIndex;

    // 起止列索引
    private int startColIndex;
    private int endColIndex;

}
