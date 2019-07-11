package com.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Bookmark;
import org.apache.poi.hwpf.usermodel.Bookmarks;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Section;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;

public class Docoperation {
	public static void main(String[] args) throws Exception {
       String path = Thread.currentThread().getContextClassLoader().getResource("templemet.doc").getPath();
       System.out.println(path);
       testWrite(path);
	}
	
	public void testReadByDoc(String path) throws Exception {
        InputStream is = new FileInputStream(path);
        @SuppressWarnings("resource")
		HWPFDocument doc = new HWPFDocument(is);
        //输出书签信息
        this.printInfo(doc.getBookmarks());
        //输出文本
        System.out.println(doc.getDocumentText());
        Range range = doc.getRange();
        this.printInfo(range);
        //读表格
        this.readTable(range);
        //读列表
        this.readList(range);
        //把当前HWPFDocument写到输出流中
        doc.write(new FileOutputStream("D:\\test.doc"));
        is.close();
    }

    /**
     * 输出书签信息
     * @param bookmarks
     */
    private void printInfo(Bookmarks bookmarks) {
        int count = bookmarks.getBookmarksCount();
        System.out.println("书签数量：" + count);
        Bookmark bookmark;
        for (int i=0; i<count; i++) {
            bookmark = bookmarks.getBookmark(i);
            System.out.println("书签" + (i+1) + "的名称是：" + bookmark.getName());
            System.out.println("开始位置：" + bookmark.getStart());
            System.out.println("结束位置：" + bookmark.getEnd());
        }
    }

    /**
     * 读表格
     * 每一个回车符代表一个段落，所以对于表格而言，每一个单元格至少包含一个段落，每行结束都是一个段落。
     * @param range
     */
    private void readTable(Range range) {
        //遍历range范围内的table。
        TableIterator tableIter = new TableIterator(range);
        Table table;
        TableRow row;
        TableCell cell;
        while (tableIter.hasNext()) {
            table = tableIter.next();
            int rowNum = table.numRows();
            for (int j=0; j<rowNum; j++) {
                row = table.getRow(j);
                int cellNum = row.numCells();
                for (int k=0; k<cellNum; k++) {
                    cell = row.getCell(k);
                    //输出单元格的文本
                    System.out.println(cell.text().trim());
                }
            }
        }
    }

    /**
     * 读列表
     * @param range
     */
    private void readList(Range range) {
        int num = range.numParagraphs();
        Paragraph para;
        for (int i=0; i<num; i++) {
            para = range.getParagraph(i);
            if (para.isInList()) {
                System.out.println("list: " + para.text());
            }
        }
    }

    /**
     * 输出Range
     * @param range
     */
    private void printInfo(Range range) {
        //获取段落数
        int paraNum = range.numParagraphs();
        System.out.println(paraNum);
        for (int i=0; i<paraNum; i++) {
            System.out.println("段落" + (i+1) + "：" + range.getParagraph(i).text());
        }
        int secNum = range.numSections();
        System.out.println(secNum);
        Section section;
        for (int i=0; i<secNum; i++) {
            section = range.getSection(i);
            System.out.println(section.getMarginLeft());
            System.out.println(section.getMarginRight());
            System.out.println(section.getMarginTop());
            System.out.println(section.getMarginBottom());
            System.out.println(section.getPageHeight());
            System.out.println(section.text());
        }
    }
    
    public static void testWrite(String templatePath) throws Exception {
        InputStream is = new FileInputStream(templatePath);
        OutputStream os = null;
        @SuppressWarnings("resource")
		HWPFDocument doc = new HWPFDocument(is);
        Range range = doc.getRange();
        for(int i=0;i<10;i++){
            //把range范围内的${reportDate}替换为当前的日期
            range.replaceText("${name}", "mike");
            os = new FileOutputStream(new File("D:\\mike.doc"));
            //把doc输出到输出流中
            doc.write(os);
        }
        os.close();
        is.close();
    }
    
}
