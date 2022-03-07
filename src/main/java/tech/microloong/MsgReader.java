package tech.microloong;

import cn.hutool.core.util.StrUtil;
import cn.hutool.json.JSONUtil;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.ExcelWriter;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.simplejavamail.outlookmessageparser.OutlookMessageParser;
import org.simplejavamail.outlookmessageparser.model.OutlookMessage;

import java.awt.*;
import java.io.File;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 读取 msg 文件
 *
 * @author MicroLOONG
 */
public class MsgReader {

    final static String WINDOWS_PATH = "C:/Users/" + System.getProperty("user.name");
    final static String MAC_PATH = ".";

    /**
     * 读取 MSG 文件名和内容
     * @param path 文件路径
     * @return 数据列表
     */
    private ArrayList<Map<String, String>> readMessage(String path) {
        // message
        ArrayList<Map<String, String>> messageList = new ArrayList<>();
        File file = new File(path + "/mail/msg");
        if (file.isDirectory()) {
            File[] files = file.listFiles();
            if (files != null) {
                try {
                    for (File temp : files) {
                        OutlookMessageParser parser = new OutlookMessageParser();
                        OutlookMessage message = parser.parseMsg(temp);
                        String time = StrUtil.subBetween(message.getBodyText(), "With effect from ", ",");
                        Map<String, String> map = new LinkedHashMap<>(4);
                        map.put("TIME", time);
                        // 获取文件名
                        String fileName = temp.getName();
                        // 前半部分
                        String[] frontPart = fileName.split("_");
                        map.put("NAME", frontPart[1].trim());
                        map.put("ID", frontPart[2].trim());
                        // 后半部分
                        String[] backPart = frontPart[3].split("-");
                        map.put("CERTIFICATE", backPart[0].trim());
                        // 末尾部分
                        String[] endPart = fileName.split("-");
                        map.put("CONTENT", StrUtil.removeSuffix(endPart[1], ".msg").trim().replace("_", "/"));
                        // append 到数组中
                        messageList.add(map);
                    }
                } catch (Exception e) {
                    System.out.println(e.getMessage());
                }
            }
        }
        // 按 NAME 排序
        messageList.sort(Comparator.comparing(o -> o.get("NAME")));

        System.out.println("结果：" + JSONUtil.parse(messageList));
        return messageList;
    }

    /**
     * Excel 导出
     * @param messageList 数据列表
     * @param path 文件路径
     */
    private void convertToExcel(ArrayList<Map<String, String>> messageList, String path) {
        // 判断目录是否存在
        File dir = new File(path + "/mail/excel");
        if (!dir.exists()) {
            dir.mkdir();
        }
        // 通过工具类创建writer
        ExcelWriter writer = ExcelUtil.getWriter(path + "/mail/excel/OUTPUT-" + DateTimeFormatter.ofPattern("yyyy-MM-dd").format(LocalDateTime.now()) +".xlsx");
        // 设置样式
        Font font = writer.createFont();
        font.setFontName("微软雅黑");
        writer.getStyleSet().setFont(font, false);
        writer.setDefaultRowHeight(20);
        writer.getStyleSet().setAlign(HorizontalAlignment.LEFT, VerticalAlignment.CENTER);

        for (int i = 0; i < writer.getColumnCount(); i++) {
            // 调整每一列宽度
            writer.getSheet().autoSizeColumn((short) i);
            // 解决自动设置列宽中文失效的问题
            writer.getSheet().setColumnWidth(i, writer.getSheet().getColumnWidth(i));
        }

        // 一次性写出内容，使用默认样式，强制输出标题
        writer.write(messageList, true);
        // 关闭writer，释放内存
        writer.close();
        System.out.println("输出完成，请查看文件");
        try {
            Desktop.getDesktop().open(new File(path + "/mail/excel"));
        } catch (IOException e) {
            System.out.println("尝试打开文件夹失败，请手动操作");
        }
    }

    public static void main(String[] args) {
        String path = null;
        if (System.getProperty("os.name").startsWith("Mac OS")) {
            path = MAC_PATH;
        } else if (System.getProperty("os.name").startsWith("Windows")) {
            path = WINDOWS_PATH;
        }
        MsgReader reader = new MsgReader();
        reader.convertToExcel(reader.readMessage(path), path);
    }
}
