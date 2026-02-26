package pers.lucas.lee.attachment2docx.utils;

import lombok.extern.slf4j.Slf4j;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Base64;

/**
 * 附件图标工具类
 *
 * @author lucas
 */
@Slf4j
public class AttachmentIconUtils {
    /**
     * 默认图表-未知文档
     */
    private static final String DEF_ICON_BASE64 = "iVBORw0KGgoAAAANSUhEUgAAAEoAAABKCAYAAAAc0MJxAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAOESURBVHhe7ZzbUhpBEIYnIB443FrkEaIG8xQWDxHfM974BPEAsbz0NsZbS7ixlmzPzqzN0rvzo8AOSX9UV4/T27Nr8/e4CPpplmIAxuM7NzJmlj6sZ6l+PEvYnDuOyONCTjrIfEoizInnEeZ2dnasJ05Ojt3ImG6360bvp+G8EkALBQK33mj8y42MOT764kb10Wi8Pcd3d9m20Om8tdjT0x83ojb8an2v9/4WVEWBwIq6HY3dKH2Gjo+sf319tb4O+MZ9f39v/eHhofXEy8uLGxnz+PvR+sHpwPper2f9MqiiQP7LQj0/P7sRDlwo6lBvsdJsNnPrdDq59T/3rY1Go9yoWMsUTFsPRAsFgrde+tLEW6zQT0JvUhv2+2kLOru5ubWGoooCge+jrq5v3MiY00F2pxvbfdTDw4P1CPv7+9YPh0PrQ6iiQLRQIHDr/by6diNjvrmXAnW23nvg7Xp5eWn92dmZ9SFUUSBaKBC4UNSh3rYN6bqXvSdURYH8s4oKXW8yS6yhqKJAtFAgeKFIwt4iJdRuc3E6ZIlvJRpFHRwcLFhMwIWid3C9rRJelOl0mhuBFqxKRYQUD+UUqVVRxQJxil/XTa2FkgrE8TFEVetmazdz3zpl7YPGUaLZzGMn6kLxPaxu4EKFpLwpqq6h7BrRuSqiVVRMaiKiVFSxSKFzS7FQzta/KI5NSZ6oChVrkYhoN/PQ+aS4NMeZi9NvN/U3nKtHCwUSVevR3jSZTCrPIV1D6LqkeGJm1lBUUSBRKCq0thSX5jiheBrIDEQVBRJVodrttrUYqbX1qtYrO9+qcqS5KqJSFP3EI4uRjSsqtI4UW0dOkt6Vk6HoZg6ihQLZWOtV5UprS3McKS7NcXi86jgJVRSIFgoELxTJ1BtISOJSrOp4YlM5RVRRIHCh0A9p+Geq7NmS4tIcR4pLc5xQPA1kBqKKAtFCgaxsM6+UeYoUr8rxsXXlLPtZL1UUiBYKBC6UJGtpjiPFpTmOFFtLTpJYQ1FFgXxIURJlx1XlSjnSHEeKS3McHvcPFFUUiBYKZGWtJ8U+klPGpnKKqKJAcEWxh4f/Q4ZWq7Vgu7u7CxaK7+3tLVgoTn+aXzQpzlFFrQktFAj85/wXPy7ciF5QujtalumX4cvxcZ7D3kvLP3Yj5PD33MQ159bOxvxOm28RPkda5/v5ufUhVFEgWigIY/4CpuWNYEfmQDwAAAAASUVORK5CYII=";
    /**
     * 默认图表-未知文档
     */
    private static final byte[] DEF_ICON_BYTES = Base64.getDecoder().decode(DEF_ICON_BASE64);

    /**
     * 根据文件名生成文件图标
     *
     * @param fileName 文件名
     * @return 图表
     */
    public static byte[] buildIconByFileName(String fileName) {
        try {
            // 先创建高分辨率图片（256x256，是目标尺寸的4倍）
            int highResWidth = 256;
            int highResHeight = 256;
            int scale = 4; // 放大倍数

            BufferedImage highResImage = new BufferedImage(highResWidth, highResHeight, BufferedImage.TYPE_INT_ARGB);
            Graphics2D g2d = highResImage.createGraphics();

            // 开启抗锯齿和高质量渲染
            g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_TEXT_ANTIALIASING, RenderingHints.VALUE_TEXT_ANTIALIAS_ON);
            g2d.setRenderingHint(RenderingHints.KEY_RENDERING, RenderingHints.VALUE_RENDER_QUALITY);
            g2d.setRenderingHint(RenderingHints.KEY_INTERPOLATION, RenderingHints.VALUE_INTERPOLATION_BICUBIC);

            // 填充白色背景
            g2d.setColor(Color.WHITE);
            g2d.fillRect(0, 0, highResWidth, highResHeight);

            // 绘制图标和文字（使用放大后的尺寸）
            String[] split = fileName.split("\\.");
            String suffix = split.length > 1 ? split[split.length - 1] : "?";
            drawIcon(suffix, g2d, highResWidth, scale);

            drawText(fileName, g2d, highResWidth, scale);

            g2d.dispose();

            // 可选：保存高分辨率版本用于对比
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            ImageIO.write(highResImage, "PNG", outputStream);
            return outputStream.toByteArray();
        } catch (IOException e) {
            log.warn("附件图标生成失败：{}", fileName, e);
        }
        return DEF_ICON_BYTES;
    }

    /**
     * 绘制图标（可以替换为加载实际图标文件）
     *
     * @param scale 放大倍数
     */
    private static void drawIcon(String iconText, Graphics2D g2d, int width, int scale) {
        // 方式1: 绘制一个简单的图标示例（紫红色圆形）
        int iconSizeX = 28 * scale;
        int iconSizeY = 28 * scale;
        int iconX = (width - iconSizeX) / 2;
        int iconY = 8 * scale;

        // 绘制圆形背景
        g2d.setColor(new Color(31, 76, 120));
        g2d.fillOval(iconX, iconY, iconSizeX, iconSizeY);

        // 绘制字母（模拟图标内容）
        g2d.setColor(Color.WHITE);
        g2d.setFont(new Font("Arial", Font.BOLD, (iconText.length() > 4 ? 8 : 12) * scale));
        FontMetrics fm = g2d.getFontMetrics();
        int textX = iconX + (iconSizeX - fm.stringWidth(iconText)) / 2;
        int textY = iconY + ((iconSizeY - fm.getHeight()) / 2) + fm.getAscent();
        g2d.drawString(iconText, textX, textY);
    }

    /**
     * 绘制文字
     *
     * @param scale 放大倍数
     */
    private static void drawText(String line, Graphics2D g2d, int width, int scale) {
        g2d.setColor(Color.BLACK);
        // 第一行文字
        Font font1 = new Font("Microsoft YaHei", Font.PLAIN, 8 * scale);
        g2d.setFont(font1);
        FontMetrics fm1 = g2d.getFontMetrics();

        char[] chars = line.toCharArray();
        int maxLine = 3;
        ArrayList<String> lines = new ArrayList<>(maxLine);
        StringBuilder currentLine = new StringBuilder();
        for (char c : chars) {
            // 检查当前行宽度
            if (fm1.stringWidth(currentLine.toString() + c) > width - 10) {
                // 当前行已满，添加到行列表
                lines.add(currentLine.toString());
                currentLine = new StringBuilder();

                // 达到两行后停止处理
                if (lines.size() >= maxLine) {
                    break;
                }
            }
            currentLine.append(c);
        }
        // 添加最后一行（如果没有达到上限）
        if (lines.size() < maxLine && currentLine.length() > 0) {
            lines.add(currentLine.toString());
        }

        // 绘制文本
        // 起始Y坐标
        int y = 42 * scale + (lines.size() == maxLine ? 0 : 15);
        // 行高
        int lineHeight = fm1.getHeight() - (lines.size() == maxLine ? 2 : 0);
        for (String l : lines) {
            g2d.drawString(l, 5, y);
            // 移动到下一行
            y += lineHeight;
        }
    }

}