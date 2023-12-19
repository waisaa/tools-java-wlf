package cn.wlftool.excel.entity;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.ToString;

@Getter
@ToString
@NoArgsConstructor
@AllArgsConstructor
public enum JavaType {

    STRING("String"),
    INTEGER("Integer"),
    LIST("List"),
    DOUBLE("Double");

    private String type;
}
