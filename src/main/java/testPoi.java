import org.apache.poi.POIXMLDocument;
import org.apache.poi.POIXMLTextExtractor;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * Created with IntelliJ IDEA.
 *
 * @author zhang
 * date: 2020/12/24 10:25
 * description:
 */
public class testPoi {
    /**
     * 读取word文件内容
     *
     * @param path
     * @return buffer
     */

    public String readWord(String path) {
        String buffer = "";
        try {
            if (path.endsWith(".doc")) {
                InputStream is = new FileInputStream(new File(path));
                WordExtractor ex = new WordExtractor(is);
                buffer = ex.getText();
                ex.close();
            } else if (path.endsWith("docx")) {
                OPCPackage opcPackage = POIXMLDocument.openPackage(path);
                POIXMLTextExtractor extractor = new XWPFWordExtractor(opcPackage);
                buffer = extractor.getText();
                extractor.close();
            } else {
                System.out.println("此文件不是word文件！");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }

        return buffer;
    }

    public static void main(String[] args) {
        // TODO Auto-generated method stub
        testPoi tp = new testPoi();
        String content = tp.readWord("C:\\Users\\zhang\\Desktop\\新建 Microsoft Word 文档.docx");
        //String arr[]=content.split("\\d+");
        //String arr[]=content.split("第"+"\\w"+"章");
        String arr[] = content.split("\\r\\n");
        /*String[] a=arr[13].split("\\d+");
        String[] b=a[1].split("\\s+");
        System.out.println(b[1]);*/
        String[] reci = new String[276];
        ;
        for (int i = 12, j = 0; i < 290; i++, j++) {
            arr[i] = arr[i] + "1";
            if (!arr[i].equals("1")) {

                if (i < 27) {//判断页面数是否为单数
                    String[] a = arr[i].split("\\d+|\\s+");
                    if (arr[i] != "\\s+") {//判断该元素是否为连续空格
                        if (a.length == 2) {//判断该元素是否为标题即分割成2个段
                            reci[j] = a[1];
                            System.out.println(a[1]);
                        } else if (a.length == 1) {
                            reci[j] = a[1];
                            System.out.println(arr[i]);
                        } else//否则该元素是平常元素可以分割成3个段
                        {
                            reci[j] = a[2];
                            System.out.println(a[2]);
                        }
                    }
                } else {
                    String[] a = arr[i].split("\\d{2,3}|\\s+|\\t");
                    if (arr[i] != "\\s+") {//判断该元素是否为连续空格
                        if (a.length == 2) {
                            reci[j] = a[1];
                            System.out.println(a[1]);
                        } else if (a.length == 1) {
                            reci[j] = a[1];
                            System.out.println(arr[i] + i);
                        } else if (a.length == 4) {
                            reci[j] = a[1];
                            System.out.println(a[1]);
                        } else {
                            reci[j] = a[2];
                            System.out.println(a[2]);
                        }
                    }
                }
            }
        }
        String fengefu = reci[0];
        for (int i = 1; i < 276; i++) {
            if (reci[i] != null)
                fengefu = fengefu + "|" + reci[i];

        }
        System.out.println(reci[275]);
        System.out.println(fengefu);
        String arr2[] = content.split(fengefu);
        for (int i = 0; i < 200; i++)
            System.out.println(arr2[i]);
    }
}
