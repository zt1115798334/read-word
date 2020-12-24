import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.CharacterRun;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.ArrayList;

/**
 * Created with IntelliJ IDEA.
 *
 * @author zhang
 * date: 2020/12/24 10:19
 * description:
 */
public class GetWord {
    public static void main(String[] args) {
        // TODO Auto-generated method stub
        try {
//            List<Policy_content> list = new ArrayList<>();
            InputStream is = new FileInputStream(new File("C:\\Users\\zhang\\Desktop\\新建 Microsoft Word 文档.doc"));  //需要将文件路更改为word文档所在路径。
            POIFSFileSystem fs = new POIFSFileSystem(is);
            HWPFDocument document = new HWPFDocument(fs);
            Range range = document.getRange();

            CharacterRun run1 = null;//用来存储第一行内容的属性
            CharacterRun run2 = null;//用来存储第二行内容的属性
            int q=1;
            for (int i = 0; i < range.numParagraphs()-1; i++) {
                Paragraph para1 = range.getParagraph(i);// 获取第i段
                Paragraph para2 = range.getParagraph(i+1);// 获取第i段
                int t=i;              //记录当前分析的段落数

                String paratext1 = para1.text().trim().replaceAll("\r\n", "");   //当前段落和下一段
                String paratext2 = para2.text().trim().replaceAll("\r\n", "");
                run1=para1.getCharacterRun(0);
                run2=para2.getCharacterRun(0);
                if (paratext1.length() > 0&&paratext2.length() > 0) {
                    //这个if语句为的是去除大标题，连续三个段落字体大小递减就跳过
                    if(run1.getFontSize()>run2.getFontSize()&&run2.getFontSize()>range.getParagraph(i+2).getCharacterRun(0).getFontSize()) {
                        continue;
                    }
                    //连续两段字体格式不同
                    if(run1.getFontSize()>run2.getFontSize()) {

                        String content=paratext2;
                        run1=run2;  //从新定位run1  run2
                        run2=range.getParagraph(t+2).getCharacterRun(0);
                        t=t+1;
                        while(run1.getFontSize()==run2.getFontSize()) {
                            //连续的相同
                            content+=range.getParagraph(t+1).text().trim().replaceAll("\r\n", "");
                            run1=run2;
                            run2=range.getParagraph(t+2).getCharacterRun(0);
                            t++;
                        }

                        if(paratext1.indexOf("HYPERLINK")==-1&&content.indexOf("HYPERLINK")==-1) {
                            System.out.println(q+"标题"+paratext1+"\t内容"+content);
                            i=t;
                            q++;
                        }

                    }
                }
            }



        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
