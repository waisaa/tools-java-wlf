package cn.wlftool.excel.constant;

import java.util.Arrays;
import java.util.List;

public class Const {

    // 模型参数文件导入路径
    public static final String MODEL_IMPORT_PATH = "model/import/";
    public static final String MODEL_IMPORT_SOIL_FILEPATH = "model/import/3_土壤性质参数.xlsx";
    public static final String MODEL_IMPORT_FILEPATH = "model/import/inputTable.xlsx";
    // 模板文件路径
    public static final String TPL_MODEL_INPUT_FILEPATH = "model/tpl/inputTable-tpl.xlsx";
    public static final String TPL_MODEL_OUTPUT_FILEPATH = "model/tpl/outputTable-tpl.xlsx";
    public static final String TPL_POINTS_FILEPATH = "model/tpl/points-tpl.xlsx";
    public static final String TPL_COMPANY_FILEPATH = "model/tpl/company-tpl.xlsx";
    public static final String TPL_POLLUTANT_EXPORT_FILEPATH = "model/tpl/pollutant-export-tpl.xlsx";
    public static final String TPL_POLLUTANT_IMPORT_FILEPATH = "model/tpl/pollutant-import-tpl.xlsx";
    // 模型计算时输入模型参数文件路径
    public static final String MODEL_INPUT_FILEPATH = "model/input/inputTable.xlsx";
    // 模型计算输出结果文件路径
    public static final String MODEL_OUTPUT_FILEPATH = "model/output/outputTable.xlsx";
    // 导入excel文件校验出错时生成的错误提示文件目录
    public static final String MODEL_CHECK_PATH = "model/check/";
    // 数据导出目录模板
    public static final String TPL_MODEL_EXPORT_FILEPATH = "model/export/output-%s.xlsx";
    public static final String TPL_EXPORT_POLLUTANT_IMPORT_FILEPATH = "model/export/pollutant-import-tpl.xlsx";

    // 默认字段
    public static final List<String> FIELDS_DEFAULT = Arrays.asList(
            "serialVersionUID", "id", "importDate", "updateDate", "updater", "estimateId");


    // format[fmt][格式]
    public static final String FMT_TRIM_SEC = "yyyyMMddHHmmss";
    public static final String FMT_STD_SEC = "yyyy-MM-dd HH:mm:ss";

    // str
    public static final String S_CN_TOXICITY = "地下水理化毒性参数";
    public static final String S_CN_EXPOSE = "受体暴露参数";
    public static final String S_CN_SOIL = "土壤性质参数";
    public static final String S_CN_WATER = "地下水性质参数";
    public static final String S_CN_BUILDING = "建筑物特征参数";
    public static final String S_CN_AIR = "空气特征参数";
    public static final String S_CN_MOVING = "离场迁移参数";
    public static final String S_DEFAULT_ESTIMATE_ID = "2023Q420231201020641";
    public static final String S_MID_EMPTY = "-| ";
    public static final String S_NULL = "null";
    public static final String S_ALL = "all";
    public static final String S_DW = "DW-";
    public static final String S_0 = "0";
    public static final String S_STATIC = "classpath:static";
    public static final String S_UTF8 = "UTF-8";
    public static final String S_OST = "超标";
    public static final String S_A0 = "A0";
    public static final String S_STD = "未超标";
    public static final String S_OS_NAME = "os.name";
    public static final String S_WINDOWS = "Windows";
    public static final String S_LINUX = "Linux";
    public static final String S_LIN_UPLOAD = "uploadFilePathUnix";
    public static final String S_WIN_UPLOAD = "uploadFilePath";
    public static final String S_TITLE = "错误提示";
    public static final String S_WATER_ENV_WARN = "水环境预警";
    public static final String S_EXCEL_LOCK_PWD = "yl@2023@excel";
    public static final String FONT_TNR = "Times New Roman";

    // 上标
    public static final String TAG_SUB_START = "<sub>";
    public static final String TAG_SUB_END = "</sub>";
    // 下标
    public static final String TAG_SUP_START = "<sup>";
    public static final String TAG_SUP_END = "</sup>";

    // punctuation[ptn][标点]
    public static final String PTN_EMPTY = "";
    public static final String PTN_BAR_MID = "-";
    public static final String PTN_SLASH = "/";

    // tpl
    public static final String TPL_TAG = "%s%s%s";
    public static final String TPL_E1 = "%s+%s";
    public static final String TPL_UPLOAD_URL = "%s/%s/%s";

    // fmt
    // 科学计数法
    public static final String FMT_DOUBLE = "0.00E00";
    public static final String FMT_SCIENTIFIC_NOTATION = "0.00E+00";
    // 文本
    public static final String FMT_TEXT = "@";

}
