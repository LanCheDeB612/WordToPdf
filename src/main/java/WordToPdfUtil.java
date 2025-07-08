import lombok.extern.slf4j.Slf4j;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.docx4j.Docx4J;
import org.docx4j.fonts.IdentityPlusMapper;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.PhysicalFonts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

import java.awt.*;
import java.io.File;
import java.io.FileOutputStream;
import java.util.Map;


/**
 * @Author LanChe
 * @Create 2025/7/8 10:00
 * Description: 将word文档转为pdf
 */
public class WordToPdfUtil {

    /**
     * 将 Word 文档转换为 PDF 文件
     * @param docxPath Word 文件路径（.docx）
     * @param pdfPath  目标 PDF 输出路径
     * @return 是否转换成功
     */
    public static boolean convertDocxToPdf(String docxPath, String pdfPath) {
        try {
            System.out.println("开始转换....");
            // 加载 Word 文档
            WordprocessingMLPackage pkg = Docx4J.load(new File(docxPath));

            // 配置字体映射
            Mapper fontMapper = new IdentityPlusMapper();
            fontMapper.put("方正小标宋简体", PhysicalFonts.get("FZXiaoBiaoSong-B05S"));
            fontMapper.put("仿宋_GB2312", PhysicalFonts.get("FangSong_GB2312"));
            fontMapper.put("方正仿宋_GBK", PhysicalFonts.get("FZFangSong-Z02"));

            pkg.setFontMapper(fontMapper);

            // 输出为 PDF
            FileOutputStream os = new FileOutputStream(pdfPath);
            Docx4J.toPDF(pkg, os);
            os.close();
            System.out.println("Word 转 PDF 成功:"+pdfPath);

            return true;
        } catch (Exception e) {
            System.err.println(" Word 转 PDF 失败：" + e.getMessage());
            e.printStackTrace();
            return false;
        }
    }
    /**
     * 打印操作系统中安装的字体（GraphicsEnvironment）
     */
    public static void printSystemFonts() {
        System.out.println("=== 操作系统可用字体列表 ===");
        GraphicsEnvironment ge = GraphicsEnvironment.getLocalGraphicsEnvironment();
        String[] fontNames = ge.getAvailableFontFamilyNames();
        for (String font : fontNames) {
            System.out.println(font);
        }
        System.out.println("共安装字体：" + fontNames.length + " 个");
    }
    /**
     * 打印系统中 docx4j 可识别的所有字体
     */
    public static void printDocx4jSystemFonts() throws Exception {
        System.out.println("开始加载系统字体...");
        PhysicalFonts.discoverPhysicalFonts();

        Map<String, PhysicalFont> fontMap = PhysicalFonts.getPhysicalFonts();
        for (Map.Entry<String, PhysicalFont> entry : fontMap.entrySet()) {
            System.out.println("Font key: " + entry.getKey() + ", Display name: " + entry.getValue().getName());
        }

        System.out.println("字体加载完毕，共 " + fontMap.size() + " 个");
    }


}
