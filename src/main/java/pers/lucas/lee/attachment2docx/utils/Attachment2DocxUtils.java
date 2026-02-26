package pers.lucas.lee.attachment2docx.utils;


import lombok.Getter;
import org.apache.poi.ooxml.POIXMLDocument;
import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ooxml.POIXMLTypeLoader;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.*;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.Ole10Native;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlObject;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.w3c.dom.Document;
import org.xml.sax.InputSource;

import java.io.*;
import java.nio.file.Files;
import java.util.List;
import java.util.UUID;
import java.util.concurrent.ThreadLocalRandom;

/**
 * docx文档中添加附件<br>
 * <pre>todo 待验证系统+软件组合是否可以打开附件，已验证系统如下：</pre>
 * 测试附件格式：doc docx xls xlsx pdf 7z zip eml<br>
 * 开发电脑1：wind10 + microsoft office + wps + adobe reader + outlook + 7z   支持以上所有<br>
 * 测试电脑1：wind11 + wps + adobe reader + 邮箱大师 + 7z                      支持以上所有<br>
 * 测试电脑2：wind11 + wps + 邮箱大师 + 7z                                     不支持打开pdf<br>
 * 测试电脑3：统信桌面操作系统(专业版1070） + wps linux + 邮箱 + 归档管理器       支持以上所有<br>
 *
 * @author lucas
 */
public class Attachment2DocxUtils {
    private static final String SHAPE_TYPE_ID = "_x0000_t79";
    private static final String SHAPE_TYPE_XML = "<v:shapetype id=\"" + SHAPE_TYPE_ID + "\" coordsize=\"21600,21600\" o:spt=\"75\" o:preferrelative=\"t\"  "
            + "                      path=\"m@4@5l@4@11@9@11@9@5xe\" filled=\"f\" stroked=\"f\">\n"
            + "                        <v:stroke joinstyle=\"miter\"/>\n"
            + "                        <v:formulas>\n"
            + "                            <v:f eqn=\"if lineDrawn pixelLineWidth 0\"/>\n"
            + "                            <v:f eqn=\"sum @0 1 0\"/>\n"
            + "                            <v:f eqn=\"sum 0 0 @1\"/>\n"
            + "                            <v:f eqn=\"prod @2 1 2\"/>\n"
            + "                            <v:f eqn=\"prod @3 21600 pixelWidth\"/>\n"
            + "                            <v:f eqn=\"prod @3 21600 pixelHeight\"/>\n"
            + "                            <v:f eqn=\"sum @0 0 1\"/>\n"
            + "                            <v:f eqn=\"prod @6 1 2\"/>\n"
            + "                            <v:f eqn=\"prod @7 21600 pixelWidth\"/>\n"
            + "                            <v:f eqn=\"sum @8 21600 0\"/>\n"
            + "                            <v:f eqn=\"prod @7 21600 pixelHeight\"/>\n"
            + "                            <v:f eqn=\"sum @10 21600 0\"/>\n"
            + "                        </v:formulas>\n"
            + "                        <v:path o:extrusionok=\"f\" gradientshapeok=\"t\" o:connecttype=\"rect\"/>\n"
            + "                        <o:lock v:ext=\"edit\" aspectratio=\"t\"/>\n"
            + "                    </v:shapetype>\n";

    /**
     * 添加多个附件到同一段落（指定段落）
     *
     * @param doc             Word 文档对象
     * @param attachmentFiles 附件文件列表
     */
    public static void addAttachmentsInSameParagraph(XWPFDocument doc, XWPFParagraph paragraph, List<Attachment> attachmentFiles) throws Exception {
        addAttachmentsInSameParagraph(doc, paragraph, attachmentFiles, null, EmbedMode.OLE_OBJECT);
    }

    /**
     * 添加多个附件到同一段落（指定模式）
     *
     * @param doc             Word 文档对象
     * @param attachmentFiles 附件文件列表
     * @param outputStream    输出文件
     * @param mode            嵌入模式
     */
    public static void addAttachmentsInSameParagraph(XWPFDocument doc, XWPFParagraph paragraph, List<Attachment> attachmentFiles,
                                                     OutputStream outputStream, EmbedMode mode) throws Exception {
        if (attachmentFiles == null || attachmentFiles.isEmpty()) {
            throw new IllegalArgumentException("附件列表不能为空");
        }
        paragraph = paragraph == null ? doc.createParagraph() : paragraph;
        if (mode == EmbedMode.DIRECT_EMBED) {
            // 直接嵌入模式：在同一段落显示多个附件名称
            XWPFRun run = paragraph.createRun();
            run.setText("附件: ");
            run.setBold(true);

            for (int i = 0; i < attachmentFiles.size(); i++) {
                Attachment attachment = attachmentFiles.get(i);
                File file = attachment.file;
                String fileName = attachment.getName();
                byte[] fileData = Files.readAllBytes(file.toPath());

                // 添加文件名
                XWPFRun fileRun = paragraph.createRun();
                fileRun.setText(fileName);

                // 添加逗号分隔（最后一个不加）
                if (i < attachmentFiles.size() - 1) {
                    fileRun.setText(", ");
                }

                // 存储文件
                addEmbedData(doc, fileData, "application/octet-stream", "/word/embeddings/" + fileName);
            }

            // 添加提示
            XWPFParagraph notePara = doc.createParagraph();
            XWPFRun noteRun = notePara.createRun();
            noteRun.setText("（提示：如果无法直接打开，可以将 .docx 文件改为 .zip 解压，在 word/embeddings/ 文件夹中找到附件）");
            noteRun.setFontSize(9);
            noteRun.setColor("808080");

        } else if (mode == EmbedMode.OLE_OBJECT) {
            // OLE 对象模式：在同一段落添加多个 OLE 对象
            for (int i = 0; i < attachmentFiles.size(); i++) {
                Attachment attachment = attachmentFiles.get(i);

                // 在同一段落添加 OLE 对象
                addOleObjectToParagraph(doc, attachment, paragraph);

                // 添加空格分隔（最后一个不加）
                if (i < attachmentFiles.size() - 1) {
                    XWPFRun spaceRun = paragraph.createRun();
                    spaceRun.setText(" "); // 空格作为间隔
                }
            }

        } else if (mode == EmbedMode.HYBRID) {
            // 混合模式：OLE 对象在同一段落，备份文件单独列出
            for (int i = 0; i < attachmentFiles.size(); i++) {
                Attachment attachment = attachmentFiles.get(i);
                addOleObjectToParagraph(doc, attachment, paragraph);

                if (i < attachmentFiles.size() - 1) {
                    XWPFRun spaceRun = paragraph.createRun();
                    spaceRun.setText(" ");
                }
            }

            // 添加备份文件说明
            XWPFParagraph backupPara = doc.createParagraph();
            XWPFRun backupRun = backupPara.createRun();
            backupRun.setText("备份附件 (WPS用户可解压docx文件提取): ");
            backupRun.setFontSize(9);
            backupRun.setColor("808080");

            for (int i = 0; i < attachmentFiles.size(); i++) {
                Attachment attachment = attachmentFiles.get(i);
                String fileName = attachment.getName();
                byte[] fileData = Files.readAllBytes(attachment.file.toPath());

                XWPFRun fileRun = backupPara.createRun();
                fileRun.setText(fileName);
                fileRun.setFontSize(9);
                fileRun.setColor("808080");

                if (i < attachmentFiles.size() - 1) {
                    fileRun.setText(", ");
                }

                addEmbedData(doc, fileData, "application/octet-stream", "/word/embeddings/backup_" + fileName);
            }
        }

        // 不保存
        if (outputStream == null) {
            return;
        }

        // 保存文件
        doc.write(outputStream);
    }

    /**
     * 添加 OLE 对象到指定段落（用于同一段落多附件）
     */
    private static void addOleObjectToParagraph(XWPFDocument doc, Attachment attachmentFile, XWPFParagraph paragraph) throws Exception {
        XWPFRun run = paragraph.createRun();
        CTR ctr = run.getCTR();

        // 读取图标图片
        byte[] image = AttachmentIconUtils.buildIconByFileName(attachmentFile.getName());
        double widthPt = pixelToPoints(64);
        double heightPt = pixelToPoints(64);
        String imageRid = doc.addPictureData(image, org.apache.poi.xwpf.usermodel.Document.PICTURE_TYPE_PNG);

        // 生成唯一 ID
        String uuidRandom = UUID.randomUUID().toString().replace("-", "") + ThreadLocalRandom.current().nextInt(1024);
        String shapeId = "_x0000_i20" + uuidRandom;

        String fileName = attachmentFile.getName().toLowerCase();
        String contentType, fileType, programId;
        boolean needsOleWrapper = false;

        if (fileName.endsWith(".docx")) {
            programId = "Word.Document.12";
            contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            fileType = ".docx";
        } else if (fileName.endsWith(".doc")) {
            programId = "Word.Document.8";
            contentType = "application/msword";
            fileType = ".doc";
        } else if (fileName.endsWith(".xlsx")) {
            programId = "Excel.Sheet.12";
            contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
            fileType = ".xlsx";
        } else if (fileName.endsWith(".xls")) {
            programId = "Excel.Sheet.8";
            contentType = "application/vnd.ms-excel";
            fileType = ".xls";
        } else {
            programId = "Package";
            contentType = "application/vnd.openxmlformats-officedocument.oleObject";
            fileType = ".bin";
            needsOleWrapper = true;
        }

        byte[] attachmentData = needsOleWrapper ? createOlePackageUsingPOI(attachmentFile.file) : Files.readAllBytes(attachmentFile.file.toPath());

        String embeddId = addEmbedData(doc, attachmentData, contentType,
                "/word/embeddings/" + uuidRandom + fileType);

        String wObjectXml = "<w:object xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\""
                + "             xmlns:v=\"urn:schemas-microsoft-com:vml\""
                + "             xmlns:o=\"urn:schemas-microsoft-com:office:office\""
                + "             xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\""
                + "             w:dxaOrig=\"1520\" w:dyaOrig=\"960\">\n" + SHAPE_TYPE_XML
                + "                    <v:shape id=\"" + shapeId + "\" type=\"#"
                + SHAPE_TYPE_ID + "\" alt=\"\" style=\"width:" + widthPt + "pt;height:" + heightPt
                + "pt;mso-width-percent:0;mso-height-percent:0;mso-width-percent:0;mso-height-percent:0\" o:ole=\"\">\n"
                + "                        <v:imagedata r:id=\"" + imageRid + "\" o:title=\"\"/>\n"
                + "                    </v:shape>\n"
                + "                    <o:OLEObject Type=\"Embed\" ProgID=\"" + programId + "\" ShapeID=\"" + shapeId
                + "\" DrawAspect=\"Icon\" ObjectID=\"" + shapeId + "\" r:id=\"" + embeddId + "\">\n"
                + "                     <o:FieldCodes>\\s</o:FieldCodes>\n"
                + "                    </o:OLEObject>"
                + "            </w:object>";

        Document document = DocumentHelper.readDocument(new InputSource(new StringReader(wObjectXml)));
        ctr.set(XmlObject.Factory.parse(document.getDocumentElement(), POIXMLTypeLoader.DEFAULT_XML_OPTIONS));
    }

    /**
     * 使用 Apache POI 自带的 Ole10Native 类创建正确的 OLE Package 对象
     */
    private static byte[] createOlePackageUsingPOI(File file) throws IOException {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();

        byte[] fileData = Files.readAllBytes(file.toPath());
        String fileName = file.getName();
        String filePath = file.getAbsolutePath();

        Ole10Native ole10 = new Ole10Native(fileName, fileName, filePath, fileData);

        try (POIFSFileSystem fs = new POIFSFileSystem()) {
            DirectoryEntry root = fs.getRoot();

            byte[] oleBytes = {1, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0};
            root.createDocument("\u0001Ole", new ByteArrayInputStream(oleBytes));

            ByteArrayOutputStream ole10Stream = new ByteArrayOutputStream();
            ole10.writeOut(ole10Stream);
            root.createDocument(Ole10Native.OLE10_NATIVE, new ByteArrayInputStream(ole10Stream.toByteArray()));

            fs.writeFilesystem(bos);
        }

        return bos.toByteArray();
    }

    private static String addEmbedData(XWPFDocument doc, byte[] embedData, String contentType, String part)
            throws InvalidFormatException {
        PackagePartName partName = PackagingURIHelper.createPartName(part);
        PackagePart packagePart = doc.getPackage().createPart(partName, contentType);

        try (OutputStream out = packagePart.getOutputStream()) {
            out.write(embedData);
        } catch (IOException e) {
            throw new POIXMLException(e);
        }
        PackageRelationship ole = doc.getPackagePart().addRelationship(partName, TargetMode.INTERNAL,
                POIXMLDocument.PACK_OBJECT_REL_TYPE);
        return ole.getId();
    }

    public static double pixelToPoints(double pixel) {
        double points = pixel * 72.0;
        points /= 96.0;
        return points;
    }

    /**
     * 附件嵌入模式
     */
    public enum EmbedMode {
        OLE_OBJECT,      // OLE 对象模式（Office 支持，WPS 可能不支持）
        DIRECT_EMBED,    // 直接嵌入模式（兼容 WPS，但需要手动解压提取）
        HYBRID           // 混合模式（同时创建 OLE 对象和直接嵌入副本，但文件是双份插入，最终文档大小受影响）
    }

    /**
     * 附件
     */
    @Getter
    public static class Attachment {
        private final String name;
        private final File file;

        public Attachment(String name, File file) {
            this.name = name;
            this.file = file;
        }

    }

}