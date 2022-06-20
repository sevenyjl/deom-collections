import cn.hutool.core.io.FileUtil;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.IOException;

/**
 * 主要学习表格的批注能力
 */
public class Demo01 {
    public static void main(String[] args) throws IOException {

        // 创建工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 创建sheet
        HSSFSheet sheet = workbook.createSheet("sheetName");
        // 创建表头
        HSSFRow header = sheet.createRow(0);
        // 表头赋值
        HSSFCell cell = null;
        String[] infos = {"序号", "考生姓名", "考生身份证号", "所属单位", "考试得分"};
        for (int i = 0; i < infos.length; i++) {
            cell = header.createCell(i);
            cell.setCellValue(infos[i]);
            HSSFCellStyle style = workbook.createCellStyle();
            HSSFFont font = workbook.createFont();
            font.setColor(Font.COLOR_RED);
            style.setFont(font);
            cell.setCellStyle(style);
            // 设置批注
            HSSFPatriarch drawingPatriarch = sheet.createDrawingPatriarch();
            HSSFComment cellComment = drawingPatriarch.createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
            // 输⼊批注信息
            cellComment.setString(new HSSFRichTextString("这是批注内容!"));
            // 添加作者,选中B5单元格,看状态栏
            cellComment.setAuthor("toad");
            cell.setCellComment(cellComment);
        }
        workbook.write(new File("E:\\aaa-project\\deom-collections\\excel-demo-poi\\src\\main\\resources\\test.xls"));
    }
}
