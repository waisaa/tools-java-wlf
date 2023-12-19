package cn.wlftool.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

@Data
@Builder
@NoArgsConstructor
@AllArgsConstructor
public class SheetCell {

    private int colIndex;
    private String content;
    private String defaultValue;
    private boolean numeric;
    private boolean scientificNotation;
    private XSSFCellStyle cellStyle;

}
