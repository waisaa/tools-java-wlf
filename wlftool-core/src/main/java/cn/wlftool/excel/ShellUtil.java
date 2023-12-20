package cn.wlftool.excel;

import cn.wlftool.excel.constant.Const;
import cn.wlftool.excel.entity.StreamGobbler;

import java.io.File;

public class ShellUtil {

    public static void main(String[] args) throws Exception {
        // windows下拷贝文件命令：copy /Y sourcefile destinationfolder
//        String cmd = "copy /Y algorithms/642183df9a1cd73bd99993g1/output.json trans/";

        // linux下拷贝文件命令：\cp sourcefile destinationfolder
        String cmd = "\\cp algorithms/642183df9a1cd73bd99993g1/output.json trans/";

        // 执行拷贝命令
        exec(cmd);
    }

    /**
     * 执行命令
     */
    public static void exec(String cmd) {
        try {
            Process proc;
            String[] commands;
            if (System.getProperty(Const.S_OS_NAME).contains(Const.S_LINUX)) {
                commands = new String[]{"/bin/sh", "-c", cmd};
                proc = Runtime.getRuntime().exec(commands);
            } else {
                proc = Runtime.getRuntime().exec(cmd);
            }
            exec(proc);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    /**
     * 执行命令
     *
     * @param cmd 执行命令
     * @param dir 执行目录
     */
    public static void exec(String cmd, String dir) {
        try {
            Process proc;
            String[] commands;
            if (System.getProperty(Const.S_OS_NAME).contains(Const.S_LINUX)) {
                commands = new String[]{"/bin/sh", "-c", cmd};
                proc = Runtime.getRuntime().exec(commands, null, new File(dir));
            } else {
                proc = Runtime.getRuntime().exec(cmd, null, new File(dir));
            }
            exec(proc);
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException(e);
        }
    }

    /**
     * 执行命令
     * 【注意此处有大坑】：
     * 1.永远要在调用waitFor()方法之前读取数据流；
     * 2.永远要先从标准错误流中读取，然后再读取标准输出流；
     */
    private static void exec(Process proc) throws Exception {
        StreamGobbler errorGobbler = new StreamGobbler(proc.getErrorStream(), "Error");
        StreamGobbler outputGobbler = new StreamGobbler(proc.getInputStream(), "Output");
        errorGobbler.start();
        outputGobbler.start();
        proc.waitFor();
    }

}
