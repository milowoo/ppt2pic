import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;

import javax.imageio.IIOImage;
import javax.imageio.ImageIO;
import javax.imageio.ImageWriter;
import javax.imageio.stream.ImageOutputStream;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.rendering.PDFRenderer;

import java.io.BufferedReader;
import java.io.Closeable;
import java.io.File;
import java.io.InputStreamReader;


/**
 * author: milowoo
 * wuchat: wucg1107
 * email: 89059049@qq.com
 */
public final class Ppt2Pic {
    //现将ppt转pdf 在将pdf转图片
    private static String pptName = "a1aa1b13f32a1c6fb2f6474941212573ad27e9e2.pptx";
    private static  String  pptPath = "/Users/wuchuangeng/Downloads/";
    private  static String tmpPath = "/Users/xxx/test/";
    private  static String desPath = "/Users/xxx/test/output/";
    private  static String softDir = "/Applications/LibreOffice.app/Contents/MacOS/";


    public static void main(String[] args) throws Exception {
        String fullName = pptPath + pptName;
        String cmd = softDir + "soffice --headless --convert-to pdf " + fullName + " --outdir " + tmpPath;

        long start = System.currentTimeMillis();
        String result = execCmd(cmd, null);
        System.out.println(result);

        String preName = getPrefixName(pptName);
        String srcPath = tmpPath + preName + ".pdf";
        pdf2png(srcPath);
        long end = System.currentTimeMillis();

        System.out.println("use time "+ (end - start));
    }

    public static String getPrefixName(String fileName){
        return fileName.substring(0,fileName.lastIndexOf("."));
    }



    public static void pdf2png(String filePath) {
        File file = new File(filePath);
        try {
            PDDocument doc = PDDocument.load(file);
            PDFRenderer renderer = new PDFRenderer(doc);
            int pageCount = doc.getNumberOfPages();
            for (int i = 0; i < pageCount; i++) {
                BufferedImage image = renderer.renderImageWithDPI(i, 95);
                OutputStream out=new FileOutputStream(new File(desPath+(i+1)+".png"));
                ImageIO.write(image, "png", out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 执行系统命令, 返回执行结果
     *
     * @param cmd 需要执行的命令
     * @param dir 执行命令的子进程的工作目录, null 表示和当前主进程工作目录相同
     */
    public static String execCmd(String cmd, File dir) throws Exception {
        StringBuilder result = new StringBuilder();

        Process process = null;
        BufferedReader bufrIn = null;
        BufferedReader bufrError = null;

        try {
            // 执行命令, 返回一个子进程对象（命令在子进程中执行）
            process = Runtime.getRuntime().exec(cmd, null, dir);

            // 方法阻塞, 等待命令执行完成（成功会返回0）
            process.waitFor();

            // 获取命令执行结果, 有两个结果: 正常的输出 和 错误的输出（PS: 子进程的输出就是主进程的输入）
            bufrIn = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
            bufrError = new BufferedReader(new InputStreamReader(process.getErrorStream(), "UTF-8"));

            // 读取输出
            String line = null;
            while ((line = bufrIn.readLine()) != null) {
                result.append(line).append('\n');
            }
            while ((line = bufrError.readLine()) != null) {
                result.append(line).append('\n');
            }

        } finally {
            closeStream(bufrIn);
            closeStream(bufrError);

            // 销毁子进程
            if (process != null) {
                process.destroy();
            }
        }

        // 返回执行结果
        return result.toString();
    }

    private static void closeStream(Closeable stream) {
        if (stream != null) {
            try {
                stream.close();
            } catch (Exception e) {
                // nothing
            }
        }
    }


}
