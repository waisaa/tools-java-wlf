package cn.wlftool;

import cn.hutool.core.lang.ConsoleTable;
import cn.hutool.core.util.ClassUtil;
import cn.hutool.core.util.StrUtil;

import java.util.Set;

/**
 * <p>
 * Wlftool是一个小的Java工具类库，让Java更方便使用。
 * </p>
 *
 * @author waisaa
 */
public class Wlftool {

    public static final String AUTHOR = "waisaa";

    private Wlftool() {
    }

    /**
     * 显示Wlftool所有的工具类
     *
     * @return 工具类名集合
     * @since 5.5.2
     */
    public static Set<Class<?>> getAllUtils() {
        return ClassUtil.scanPackage("cn.wlftool",
                (clazz) -> (!clazz.isInterface()) && StrUtil.endWith(clazz.getSimpleName(), "Util"));
    }

    /**
     * 控制台打印所有工具类
     */
    public static void printAllUtils() {
        final Set<Class<?>> allUtils = getAllUtils();
        final ConsoleTable consoleTable = ConsoleTable.create().addHeader("工具类名", "所在包");
        for (Class<?> clazz : allUtils) {
            consoleTable.addBody(clazz.getSimpleName(), clazz.getPackage().getName());
        }
        consoleTable.print();
    }
}
