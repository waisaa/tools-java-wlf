package cn.wlftool.excel.entity;

import com.yunlu.groundwater.gwParameters.entities.*;
import com.yunlu.groundwater.onepage.entities.CompanyInfo;
import com.yunlu.groundwater.onepage.entities.PointInfo;
import com.yunlu.groundwater.onepage.entities.PollutantImport;
import com.yunlu.groundwater.riskManagement.entity.*;

import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

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

    // excel模型输入参数文件，各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_INPUT_SHEET_OBJ = new HashMap<String, Class<?>>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("3_地下水理化毒性报表", GWBPhysicalChemicalToxicity.class);
            put("4_受体暴露参数", GWBReceptorExpose.class);
            put("5_土壤性质参数", GWBSoilNature.class);
            put("6_地下水性质参数", GWBWaterNature.class);
            put("7_建筑物特征参数", GWBBuildingFeature.class);
            put("8_空气特征参数", GWBAirFeature.class);
            put("9_离场迁移参数", GWBFieldMoving.class);
        }
    };
    public static final Map<String, Integer> MODEL_INPUT_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("3_地下水理化毒性报表", 3);
            put("4_受体暴露参数", 3);
            put("5_土壤性质参数", 3);
            put("6_地下水性质参数", 3);
            put("7_建筑物特征参数", 3);
            put("8_空气特征参数", 3);
            put("9_离场迁移参数", 3);
        }
    };

    // excel模型输入参数文件，各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_INPUT_SHEET_OBJ_RUNNING = new HashMap<String, Class<?>>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
        	put("1_地下水点位-污染物浓度", PointSummary.class);
            put("2_点位侧向迁移距离", GWBPointMoving.class);
            put("3_地下水理化毒性报表", GWBPhysicalChemicalToxicity.class);
            put("4_受体暴露参数", GWBReceptorExpose.class);
            put("5_土壤性质参数", GWBSoilNature.class);
            put("6_地下水性质参数", GWBWaterNature.class);
            put("7_建筑物特征参数", GWBBuildingFeature.class);
            put("8_空气特征参数", GWBAirFeature.class);
            put("9_离场迁移参数", GWBFieldMoving.class);
        }
    };
    public static final Map<String, Integer> MODEL_INPUT_SHEET_START_ROW_INDEX_RUNNING = new HashMap<String, Integer>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
        	put("1_地下水点位-污染物浓度", 7);
            put("2_点位侧向迁移距离", 1);
            put("3_地下水理化毒性报表", 3);
            put("4_受体暴露参数", 3);
            put("5_土壤性质参数", 3);
            put("6_地下水性质参数", 3);
            put("7_建筑物特征参数", 3);
            put("8_空气特征参数", 3);
            put("9_离场迁移参数", 3);
        }
    };

    // excel模型数据结果文件，各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_OUTPUT_SHEET_OBJ = new HashMap<String, Class<?>>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("1_各污染物控制值 筛选值", ResultRiskControl.class);
            put("2_地下水暴露量", ResultWaterExpose.class);
            put("3_致癌风险", ResultCancerRiskExcel.class);
            put("4_非致癌危害", ResultNonCancerRiskExcel.class);
            put("5_暴露途径贡献率致癌", ResultExposePathwayCancer.class);
            put("6_暴露途径贡献率非致癌", ResultExposePathwayNonCancer.class);
            put("7_过程因子", RiskProcessFactor.class);
            put("8_介质浓度", RiskMediumConcentration.class);
            put("9_污染物特征统计值", RiskPollutantFeature.class);
            put("10_点位超过四类水污染物侧向迁移浓度变化", RiskDistanceChange.class);
            put("11_点位超过四类水污染物随时间浓度变化", RiskTimeChange.class);
        }
    };
    public static final Map<String, Integer> MODEL_OUTPUT_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("1_各污染物控制值 筛选值", 4);
            put("2_地下水暴露量", 3);
            put("3_致癌风险", 2);
            put("4_非致癌危害", 2);
            put("5_暴露途径贡献率致癌", 3);
            put("6_暴露途径贡献率非致癌", 4);
            put("7_过程因子", 2);
            put("8_介质浓度", 4);
            put("9_污染物特征统计值", 1);
            put("10_点位超过四类水污染物侧向迁移浓度变化", 1);
            put("11_点位超过四类水污染物随时间浓度变化", 1);
        }
    };

    // 各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_OUTPUT_EXPORT_SHEET_OBJ = new HashMap<String, Class<?>>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("控制值-筛选值", ResultRiskControl.class);
            put("污染物特征值", RiskPollutantFeature.class);
            put("地下水暴露量", ResultWaterExpose.class);
            put("致癌风险", PollutantSummary.class);
            put("非致癌危害", PollutantSummary.class);
            put("贡献率（致癌）", ResultExposePathwayCancer.class);
            put("贡献率（非致癌）", ResultExposePathwayNonCancer.class);
            put("过程因子", RiskProcessFactor.class);
            put("介质浓度", RiskMediumConcentration.class);
        }
    };
    // 各个sheet对应的起始行索引
    public static final Map<String, Integer> MODEL_OUTPUT_EXPORT_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("控制值-筛选值", 1);
            put("污染物特征值", 1);
            put("地下水暴露量", 2);
            put("致癌风险", 1);
            put("非致癌危害", 1);
            put("贡献率（致癌）", 2);
            put("贡献率（非致癌）", 2);
            put("过程因子", 2);
            put("介质浓度", 1);
        }
    };

    // 各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_POINTS_SHEET_OBJ = new HashMap<String, Class<?>>() {
        {
            put("监测点位管理", PointInfo.class);
        }
    };
    // 各个sheet对应的起始行索引
    public static final Map<String, Integer> MODEL_POINTS_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        {
            put("监测点位管理", 1);
        }
    };

    // 各个sheet对应的数据存储对象
    public static final Map<String, Class<?>> MODEL_COMPANY_SHEET_OBJ = new HashMap<String, Class<?>>() {
        {
            put("企业管理", CompanyInfo.class);
        }
    };
    // 各个sheet对应的起始行索引
    public static final Map<String, Integer> MODEL_COMPANY_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        {
            put("企业管理", 1);
        }
    };

    public static final Map<String, Class<?>> POLLUTANT_IMPORT_SHEET_OBJ = new HashMap<String, Class<?>>() {
        {
            put("0", PollutantImport.class);
        }
    };

    public static final Map<String, Integer> POLLUTANT_IMPORT_SHEET_START_ROW_INDEX = new HashMap<String, Integer>() {
        {
            put("0", 3);
        }
    };

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
    public static final String S_WINDOWS = "windows";
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

    public static final Map<String, String> sign_map = new HashMap<String, String>() {
        /**
         *
         */
        private static final long serialVersionUID = 1L;

        {
            put("ρ<sub>b</sub><sup>a</sup>", "\\sideset{}{_{b}^a}ρ");
            put("θ<sub>e</sub>", "_{θ_{e}}");
            put("f<sub>oc</sub><sup>a</sup>", "\\sideset{}{_{oc}^a}f");
            put("δ<sub>gw</sub>", "_{δ_{gw}}");
            put("L<sub>gw</sub>", "_{L_{gw}}");
            put("F<sub>(x)</sub>", "_{F_{(x)}}");
            put("P<sub>e</sub>", "_{P_{e}}");
            put("PM<sub>10</sub>", "_{PM_{10}}");
            put("u<sub>t</sub>", "_{u_{t}}");
            put("δ<sub>air</sub>", "_{δ_{air}}");
            put("U<sub>air</sub>", "_{U_{air}}");
            put("g·m<sup>-2</sup>·s<sup>-1</sup>/kg·m<sup>-3</sup>", "g/m^{-2}·s^{-1}/kg·m^{-3}");
            put("g/m<sup>2</sup>·s<sup>-1</sup>", "g/m^{2}·s^{-1}");
            put("mg/m<sup>3</sup>", "mg/m^{3}");
            put("L<sub>crack</sub>", "_{L_{crack}}");
            put("X<sub>crack</sub>", "_{X_{crack}}");
            put("Z<sub>crack</sub>", "_{Z_{crack}}");
            put("θ<sub>wcrack</sub>", "_{θ_{wcrack}}");
            put("L<sub>B</sub>", "_{L_{B}}");
            put("A<sub>b</sub>", "_{A_{b}}");
            put("θ<sub>acrack</sub>", "_{θ_{acrack}}");
            put("y<sub>air</sub>", "_{y_{air}}");
            put("z<sub>air</sub>", "_{z_{air}}");
            put("S<sub>d</sub>", "_{S_{d}}");
            put("S<sub>w</sub>", "_{S_{w}}");
            put("AT<sub>ca</sub>", "_{AT_{ca}}");
            put("GWCR<sub>a</sub>", "_{GWCR_{a}}");
            put("EFI<sub>a</sub>", "_{EFI_{a}}");
            put("AT<sub>nc</sub>", "_{AT_{nc}}");
            put("ED<sub>a</sub>", "_{ED_{a}}");
            put("DAIR<sub>a</sub>", "_{DAIR_{a}}");
            put("SER<sub>a</sub>", "_{SER_{a}}");
            put("EF<sub>a</sub>", "_{EF_{a}}");
            put("BW<sub>a</sub>", "_{BW_{a}}");
            put("H<sub>a</sub>", "_{H_{a}}");
            put("EFO<sub>a</sub>", "_{EFO_{a}}");
            put("m<sup>3</sup>/d", "m^{3}/d");
            put("K<sub>v</sub>", "_{K_{v}}");
            put("θ<sub>ws</sub>", "_{θ_{ws}}");
            put("ρ<sub>b</sub>", "_{ρ_{b}}");
            put("θ<sub>wcap</sub>", "_{θ_{wcap}}");
            put("h<sub>cap</sub>", "_{h_{cap}}");
            put("θ<sub>as</sub>", "_{θ_{as}}");
            put("θ<sub>acap</sub>", "_{θ_{acap}}");
            put("f<sub>oc</sub>", "_{f_{oc}}");
            put("m<sup>2</sup>", "m^{2}");
            put("g/cm<sup>3</sup>", "g/cm^{3}");
        }
    };
}
