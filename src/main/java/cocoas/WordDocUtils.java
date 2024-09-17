package cocoas;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;

public class WordDocUtils {

    /**
     * 设置纸张大小
     *
     * @param document doc对象
     * @param width    宽
     * @param height   长
     */
    public static void setPageSize(XWPFDocument document, long width, long height) {
        CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
        CTPageSz pgsz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pgsz.setW(BigInteger.valueOf(width));
        pgsz.setH(BigInteger.valueOf(height));
    }

    /**
     * 保存文件
     *
     * @param document doc对象
     * @param savePath 保存路径
     * @param fileName 文件名称
     */
    public static void saveDoc(XWPFDocument document, String savePath, String fileName) throws IOException {
        File file = new File(savePath);
        if (!file.exists()) {
            // 判断生成目录是否存在，不存在时创建目录。
            file.mkdirs();
        }
        // 保存
        fileName += ".docx";
        FileOutputStream out = new FileOutputStream(new File(savePath + File.separator + fileName));
        document.write(out);
        // 关闭资源
        out.flush();
        out.close();
        document.close();
    }


}
