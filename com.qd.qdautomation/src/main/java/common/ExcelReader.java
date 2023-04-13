package common;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.List;

public class ExcelReader {
    /*
    *完成对excel文件的读取操作\
    * @method ExcelReader构造函数，读取excel文件内容到workbook中，usesheet读取指定sheet页
    * close完成文件读取，释放资源readNextLine读取当前行，并将焦点移动到下一行readLine读取指定行
    * getCellValue针对单元格内容不同格式进行读取
    *
    * */
//xlsx格式
    private Workbook workbook;
//sheet页
    private Sheet sheet;
    //最大行数
    public int rows;
//    创建对象时，打开excel
    public ExcelReader(String path){
        //截取.后面的后缀名
        String type=path.substring(path.lastIndexOf("."));
//        初始化文件流
        FileInputStream in=null;
        try {
//            通过文件流打开文件
            in=new FileInputStream(new File(path));
            if(type.equals(".xlsx")){
                workbook=new XSSFWorkbook(in);
                sheet=workbook.getSheetAt(0);
                rows=sheet.getPhysicalNumberOfRows();
            }else if(type.equals(".xls")){
                workbook=new XSSFWorkbook(in);
                sheet=workbook.getSheetAt(0);
                rows=sheet.getPhysicalNumberOfRows();
            }
        } catch (Exception e) {
            //读取失败则给出Excel读取失败的提示，并停止
            e.printStackTrace();
            System.out.println("文件读取失败");
//            停止
            return;
        }
        try {
//            完毕后，关闭文件流
            in.close();
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
    // 获取当前Excel的所有sheet页的个数
    public int getTotalSheetNo() {
        int sheets = 0;
        if (workbook != null) {
            sheets = workbook.getNumberOfSheets();
        }
        return sheets;
    }
    //获取当前sheet页中最大行数
    public int getRowNo(){
        return rows;
    }
//获取当前sheet页的名字
    public String getSheetName(int sheetIndex){
        String sheetname="";
        if(workbook!=null){
            sheetname=workbook.getSheetName(sheetIndex);
        }
        return sheetname;
    }
//根据sheet序号指定使用的sheet页
    public void useSheetByIndex(int sheetIndex){
        if(workbook!=null){
            Sheet sheetAt = workbook.getSheetAt(sheetIndex);
            rows=sheetAt.getPhysicalNumberOfRows();
        }else {
            System.out.println("未打开Excel文件");
        }
    }
    //根据sheetName来使用sheet
    public void useSheetByName(String sheetName){
        if (workbook!=null) {
            Sheet sheet = workbook.getSheet(sheetName);
            rows= sheet.getPhysicalNumberOfRows();
        }else {
            System.out.println("未打开Excel文件");
        }
    }

//读取参数中的指定行
    public List<String> readLine(int rowNo){
        List<String> line=new ArrayList<>();
        Row row = sheet.getRow(rowNo);
        int numberOfCells = row.getPhysicalNumberOfCells();
        for (int i = 0; i < numberOfCells; i++) {
            String cellValue = getCellValue(row.getCell(i));
            line.add(cellValue);
        }
        return line;
    }
//    读取参数中的指定列
    public List<String> readColumn(int colNo){
        List<String> column=new ArrayList<>();
        for (int i = 0; i < rows; i++) {
            Row row = sheet.getRow(i);
            column.add(getCellValue(row.getCell(colNo)));
        }
        return column;
    }
//    读取指定单元格
    public String readCell(int rowNo,int colNo){
        Row row = sheet.getRow(rowNo);
        String cellValue = getCellValue(row.getCell(colNo));
        return cellValue;
    }
//以二维数组形式读取excel文件内容
public Object[][] readAsMatrix() {
    //获取当前sheet页中第一行的最大单元格数。
    int cellcount = sheet.getRow(0).getPhysicalNumberOfCells();
    //二维数组的下标，由excel的最大行数决定，以及最大列数决定。
    Object[][] matrix = new Object[rows - 1][cellcount];
    //用例从excel中的第2行开始读取，遍历到最后一行
    for (int rowNo = 1; rowNo < rows; rowNo++) {
        //遍历行中所有的单元格
        for (int colNo = 0; colNo < cellcount; colNo++) {
            matrix[rowNo - 1][colNo] = readCell(rowNo, colNo);
        }
    }
    //完成循环之后，二维数组已经存储好了对应的值，返回该二维数组。
    return matrix;
}
//    针对单元格内容不同格式进行读取
public String getCellValue(Cell cell){
        String cellValue="";
//        如果对象为null，则可能是xls文件转xlsx文件格式问题导致读取空单元格时，读到null
    if(cell==null){
        return cellValue;
    }
    //获取单元格类型。
    try {
    CellType cellType = cell.getCellType();
    // 将所有格式转为字符串读取到cellValue
    switch (cellType) {
        case STRING: // 文本
            cellValue = cell.getStringCellValue();
            break;
        case NUMERIC: // 数字、日期
            if (DateUtil.isCellDateFormatted(cell)) {
                //日期型以年-月-日格式存储
                SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
                cellValue = fmt.format(cell.getDateCellValue());
            } else {
                //数字保留两位小数
                Double d = cell.getNumericCellValue();
                DecimalFormat df = new DecimalFormat("#.##");
                cellValue = df.format(d);
            }
            break;
        case BOOLEAN: // 布尔型
            cellValue = String.valueOf(cell.getBooleanCellValue());
            break;
        case BLANK: // 空白
            cellValue = cell.getStringCellValue();
            break;
        case ERROR: // 错误
            cellValue = "错误";
            break;
        case FORMULA: // 公式
            FormulaEvaluator eval;
            eval = workbook.getCreationHelper().createFormulaEvaluator();
            cellValue = getCellValue(eval.evaluateInCell(cell));
            break;
        case _NONE:
            cellValue = "";
        default:
            cellValue = "错误";
    }
} catch (Exception e) {
        e.printStackTrace();
    }
        return cellValue;
}



    // 读取完成，关闭Excel
    public void close() {
        try {
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
