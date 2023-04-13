package common;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

public class ExcelWriter { /*
 * 读取用例表内容，并复制用例表内容到结果表中完成结果表的创建，并实现向结果表中写入数据
 *
 * @method excelWriter构造函数完成对结果表的创建 setStyle指定向结果表写入内容时的格式
 * writeFailCell用例失败时以红色字体写入 writeCell正常情况下以黑色字体写入单元格 useSheet指定使用的sheet
 * writeLine整行写入内容 save保存写入的内容到excel文件中，完成操作后必须调用，否则excel文件无内容。
 */

    // 用于存放结果文件的workbook
    private Workbook resultworkbook = null;
    // 工作的sheet页
    private Sheet sheet;
    // 用于读取用例表内容复制到结果标的文件输出流
    private FileOutputStream stream = null;
    // 用于存储结果文件的路径的成员变量，便于在保存结果时调用
    private String path = null;
    // 单元格格式
    private CellStyle style = null;
    // sheet页的总行数
    private int rows = 0;

    /*
     * 根据用例表path1，创建结果表path2，将path1中的内容复制到path2中
     *
     * @param path1用例表路径 path2结果表路径
     */
    public ExcelWriter(String caseFilePath, String resultFileName) {
        //将excel文件中的内容读取到内存中作为read工作簿对象
        // 截取用例表后缀名
        String Origintype = caseFilePath.substring(caseFilePath.lastIndexOf("."));
        // 判断是xls还是xlsx格式，完成在内存中创建用例表的工作簿
        Workbook casebook = null;
        try {
            if (Origintype.equals(".xlsx")) {
                casebook = new XSSFWorkbook(new File(caseFilePath));
            } else if (Origintype.equals(".xls")) {
                casebook = new HSSFWorkbook(new FileInputStream(new File(caseFilePath)));
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 如果两种格式均不符合，则文件打开失败，停止方法执行。
        if (casebook == null) {
            System.out.println("Excel用例文件打开失败！");
            return;
        }

        //复制read工作簿中的内容到结果文件中去。
        //拼接结果文件名
        String resultFilepath = resultFileName + Origintype;
        try {
            // 根据结果表的文件名，为该文件在内存中开辟空间
            File resultfile = new File(resultFilepath);
            try {
                // 在磁盘上面创建该文件
                resultfile.createNewFile();
            } catch (Exception e1) {
                // 创建失败，提示路径非法，并停止方法执行。
                System.out.println("结果文件路径非法，请检查！");
                e1.printStackTrace();
                return;
            }
            // 基于结果表，创建文件输出流stream
            stream = new FileOutputStream(resultfile);
            // 将用例表中的内容写入文件输出流stream
            casebook.write(stream);
            // 关闭已经写入了用例表内容的文件流
            stream.close();

            // 基于结果表，创建文件输入流，打开结果表，所有的后续写入操作都基于结果表进行操作，
            //成员变量现在的对象是结果表中的内容
            FileInputStream in = new FileInputStream(resultfile);
            // 判断结果文件的后缀是03版还是07版excel
            try {
                if (Origintype.equals(".xlsx")) {
                    resultworkbook = new XSSFWorkbook(in);
                } else if (Origintype.equals(".xls")) {
                    resultworkbook = new HSSFWorkbook(in);
                }
                // 初始化sheet页
                sheet = resultworkbook.getSheetAt(0);
                // 获取最大行数
                rows = sheet.getPhysicalNumberOfRows();
            } catch (Exception e) {
                e.printStackTrace();
            }
            // 关闭文件输入流
            in.close();
            // 将成员变量结果文件路径赋值为path2，表示结果表已经成功创建。
            path = resultFilepath;
            // 设置默认样式为第一个单元格的样式
            setStyle(0, 0);
            // System.out.println(path);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 获取当前Excel的所有sheet页的个数
    public int getTotalSheetNo() {
        int sheets = 0;
        if (resultworkbook != null) {
            sheets = resultworkbook.getNumberOfSheets();
        }
        return sheets;
    }

    //获取当前sheet页中最大行数。
    public int getRowNo() {
        return rows;
    }

    // 获取当前sheet页面的名字
    public String getSheetName(int sheetIndex) {
        String sheetname = "";
        if (resultworkbook != null)
            sheetname = resultworkbook.getSheetName(sheetIndex);
        return sheetname;
    }

    // 根据sheet名使用sheet页。
    public void useSheet(String sheetName) {
        if (resultworkbook != null) {
            // 根据打开的工作簿类型，在其中初始化sheet页
            sheet=resultworkbook.getSheet(sheetName);
            rows = sheet.getPhysicalNumberOfRows();
        } else {
            System.out.println("未打开Excel文件！");
        }
    }

    //根据sheet序号指定使用的sheet
    public void useSheetByIndex(int sheetIndex) {
        if (resultworkbook != null) {
            sheet=resultworkbook.getSheetAt(sheetIndex);
            rows = sheet.getPhysicalNumberOfRows();
        } else {
            System.out.println("error::未打开Excel文件！");
        }
    }

    // 设置样式为Excel中指定单元格的样式
    public void setStyle(int rowNo, int column) {
        Row row = null;
        Cell cell = null;
        try {
            // 获取指定行
            row = sheet.getRow(rowNo);
            // 获取指定列
            cell = row.getCell(column);
            // System.out.println(cell.getStringCellValue());
            // 保存指定单元格样式
            style = cell.getCellStyle();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 以默认格式向单元格写入内容，参数使用与writeFailCell一致
    public void writeCell(int rowNo, int column, String value) {
        Row row = null;
        try {
            // 获取指定行
            row = sheet.getRow(rowNo);
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 行不存在，则创建
        if (row == null) {
            row = sheet.createRow(rowNo);
        }
        // 在该行，新建指定列的单元格
        Cell cell = row.createCell(column);
        // 设置单元格值
        cell.setCellValue(value);
        // 设置单元格样式
        cell.setCellStyle(style);
    }

    /*
     * 当用例执行结果失败时，使用该方法，以红色字体写入excel
     *
     * @param r单元格行数 l单元格列数 value输入值
     */
    public void writeFailCell(int rowNo, int column, String value) {
        Row row = null;
        try {
            // 获取指定行
            row = sheet.getRow(rowNo);
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 行不存在，则创建
        if (row == null) {
            row = sheet.createRow(rowNo);
        }
        // 在该行，新建指定列的单元格
        Cell cell = row.createCell(column);
        // 设置单元格值
        cell.setCellValue(value);
        // 设置单元格样式
        CellStyle failStyle = null;
        // 新建字体样式
        Font font = null;
        // 根据不同的excel版本进行实例化
        failStyle = resultworkbook.createCellStyle();
        // 设置字体颜色为红色
        font = resultworkbook.createFont();
        font.setColor(IndexedColors.RED.index);
        // 将字体颜色作为单元格样式
        failStyle.setFont(font);
        //设置单元格边框
        failStyle.setBorderLeft(BorderStyle.THIN);
        failStyle.setBorderRight(BorderStyle.THIN);
        failStyle.setBorderTop(BorderStyle.THIN);
        failStyle.setBorderBottom(BorderStyle.THIN);
        // 设置对应单元格样式
        cell.setCellStyle(failStyle);
    }

    // 写入一整行的内容
    public void writeLine(int rowNo, List<String> list) {
        Row row = null;
        try {
            // 获取指定行
            row = sheet.getRow(rowNo);
        } catch (Exception e) {
            e.printStackTrace();
        }
        // 行不存在，则创建
        if (row == null) {
            row = sheet.createRow(rowNo);
        }
        Cell cell = null;
        for (int i = 0; i < list.size(); i++) {
            // 在该行，新建指定列的单元格
            cell = row.createCell(i);
            // 设置单元格值
            cell.setCellValue(list.get(i));
            // 设置单元格样式
            cell.setCellStyle(style);
        }
    }

    // 将结果表在内存中的工作簿内容保存到磁盘文件中
    public void save() {
        // 如果结果表文件未创建，则不保存
        if (path != null) {
            try {
                // 基于结果表路径创建文件输出流
                stream = new FileOutputStream(new File(path));
                // 将结果表的workbook工作簿的内容写入输出流中，即写入文件
                if (resultworkbook != null) {
                    resultworkbook.write(stream);
                    resultworkbook.close();
                } else {
                    System.out.println("未打开Excel文件！");
                }
                // 关闭输出流，完成将内存中workbook写入文件的过程，保存文件。
                stream.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
