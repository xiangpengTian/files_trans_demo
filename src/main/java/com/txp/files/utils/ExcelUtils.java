package com.txp.files.utils;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

@Slf4j
public class ExcelUtils {
    private final static String XLS = "xls";
    private final static String XLSX = "xlsx";

    public static List<String[]> readExcel(MultipartFile formFile) throws IOException {
        //检查文件
        checkFile(formFile);
        //获得工作簿对象
        Workbook workbook = getWorkBook(formFile);
        //创建返回对象，把每行中的值作为一个数组，所有的行作为一个集合返回
        List<String[]> list = new ArrayList<>();
        if (null != workbook) {
            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                //获取当前sheet工作表
                Sheet sheet = workbook.getSheetAt(sheetNum);
                if (null == sheet) {
                    continue;
                }
                //获得当前sheet的开始行
                int firstRowNum = sheet.getFirstRowNum();
                //获得当前sheet的结束行
                int lastRowNum = sheet.getLastRowNum();
                //循环除了第一行之外的所有行
                for (int rowNum = firstRowNum + 1; rowNum <= lastRowNum; rowNum++) {
                    //获得当前行
                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    //后的当前行的开始列
                    int firstCellNum = row.getFirstCellNum();
                    //获得当前行的列数
                    int lastCellNum = row.getPhysicalNumberOfCells();
                    String[] cells = new String[row.getPhysicalNumberOfCells()];
                    //循环当前行
                    for (int cellNum = firstCellNum; cellNum < lastCellNum; cellNum++) {
                        Cell cell = row.getCell(cellNum);
                        cells[cellNum] = getCellValue(cell);
                    }
                    list.add(cells);
                }
            }
        }
        return list;
    }

    /**
     * 获取当前行数据
     *
     * @param cell
     * @return
     */
    @SuppressWarnings("deprecation")
    private static String getCellValue(Cell cell) {
        String cellValue = "";

        if (cell == null) {
            return cellValue;
        }
        //把数字当成String来读，避免出现1读成1.0的情况
        if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
            cell.setCellType(Cell.CELL_TYPE_STRING);
        }
        //判断数据的类型
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC://数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING://字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN://Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case Cell.CELL_TYPE_FORMULA://公式
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case Cell.CELL_TYPE_BLANK://空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR://故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }


    /**
     * 获得工作簿对象
     *
     * @param formFile
     * @return
     */
    public static Workbook getWorkBook(MultipartFile formFile) {
        //获得文件名
//		String fileName = formFile.getName();
        String fileName = formFile.getOriginalFilename();
        //创建Workbook工作簿对象，表示整个excel
        Workbook workbook = null;
        try {
            //获得excel文件的io流
            InputStream is = formFile.getInputStream();
            //根据文件后缀名不同（xls和xlsx）获得不同的workbook实现类对象
            if (fileName.endsWith(XLS)) {
                //2003
                workbook = new HSSFWorkbook(is);
            } else if (fileName.endsWith(XLSX)) {
                //2007
                workbook = new XSSFWorkbook(is);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook;
    }

    /**
     * 检查文件 
     *
     * @param formFile
     * @throws IOException
     */
    public static void checkFile(MultipartFile formFile) throws IOException {
        //判断文件是否存在
        if (null == formFile) {
            log.error("文件不存在！");
            throw new FileNotFoundException("文件不存在！");
        }
        //获得文件名
//		String fileName = formFile.getName();
        String fileName = formFile.getOriginalFilename();
        //判断文件是否是excel文件
        if (!fileName.endsWith(XLS) && !fileName.endsWith(XLSX)) {
            log.error(fileName + "不是excel文件！");
            throw new IOException(fileName + "不是excel文件！");
        }
    }


    //java实体类 生成excel
    public static void createWorkbook(List<T> list,String fileName,String path) throws IOException {
        //创建一个workbook,对应一个excel
        HSSFWorkbook wb = new HSSFWorkbook();

        //生成单元格样式
        //创建字体样式
        HSSFFont font = wb.createFont();
        //设置字体
        font.setFontName("宋体");
        //设置字体大小
        font.setFontHeightInPoints((short) 10);
        //设置加粗
        font.setBold(true);

        //然后创建单元格样式style
        HSSFCellStyle style1 = wb.createCellStyle();
        //将字体注入
        style1.setFont(font);
        // 自动换行
        style1.setWrapText(true);
        // 水平居中
        style1.setAlignment(HorizontalAlignment.CENTER);
        // 上下居中
        style1.setVerticalAlignment(VerticalAlignment.CENTER);
        // 设置单元格的背景颜色
        style1.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());
        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        // 边框的大小
        style1.setBorderTop(BorderStyle.THIN);
        style1.setBorderBottom(BorderStyle.THIN);
        style1.setBorderLeft(BorderStyle.THIN);
        style1.setBorderRight(BorderStyle.THIN);

        //2、生成一个sheet，对应excel的sheet，参数为excel中sheet显示的名字
        HSSFSheet sheet = wb.createSheet("采集对象一致率");

        //3、设置sheet中每列的宽度，第一个参数为第几列，0为第一列；第二个参数为列的宽度，可以设置为0。
        // Test中有三个属性，因此这里设置三列，第0列设置宽度为0，第1~3列用以存放数据
        sheet.setColumnWidth(0, 0);
        sheet.setColumnWidth(1, 20 * 256);
        sheet.setColumnWidth(2, 20 * 256);
        sheet.setColumnWidth(3, 20 * 256);


        //4、生成sheet中一行，从0开始
        HSSFRow row = sheet.createRow(0);

        // 设定行的高度
        row.setHeight((short) 800);
        //5、创建row中的单元格，从0开始
        //我们第一列设置宽度为0，不会显示，因此第0个单元格不需要设置样式
        HSSFCell cell = row.createCell(0);
        //从第1个单元格开始，设置每个单元格样式
        cell = row.createCell(1);
        //设置单元格中内容
        cell.setCellValue("x");
        //设置单元格样式
        cell.setCellStyle(style1);
        //第二个单元格
        cell = row.createCell(2);
        cell.setCellValue("y");
        cell.setCellStyle(style1);
        //第三个单元格
        cell = row.createCell(3);
        cell.setCellValue("value");
        cell.setCellStyle(style1);

        //设置冻结
        //要冻结的列数、要冻结的行数、右边区域可见的左边列数，下面区域可见的首行。
        //常用于对表格表头冻结，下列数据滚动表头仍保持在上方。
        sheet.createFreezePane(0, 1,0, 1);

        //6、输入数据
     /*   for (int i = 1; i <= list.size(); i++) {
            cell = row.createCell(i);
            //操作同第5步，通过setCellValue(list.get(i-1).getX())注入数据

        }*/
        //7、单元格合并
        sheet.addMergedRegion(new CellRangeAddress(2, 3, 1, 1));//参数为(第一行，最后一行，第一列，最后一列)
        //8、输入excel
        FileOutputStream os = new FileOutputStream(path+fileName);
        wb.write(os);
        os.close();
    }


    /**
     * cell.setEncoding(HSSFCell.ENCODING_UTF_16);
     * 设置cell中格式为utf-16.用以应付可能存在的中文乱码问题。
     *
     * HSSFCellStyle style = workbook.createCellStyle();//首先创建一个style
     *
     * style.setAlignment(align);//水平居左、居右、居中。使用HSSFCellStyle自带的参数：HSSFCellStyle.ALIGN_CENTER or ALIGN_LEFT or ALIGN_RIGHT
     *
     * style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居左、居右、居中。使用HSSFCellStyle自带的参数：HSSFCellStyle.VERTICAL_CENTER or VERTICAL_LEFT or VERTICAL_RIGHT
     *
     * style.setWrapText(true);//自动换行
     *
     * HSSFFont hssfFont = workbook.createFont();//创建字体
     * hssfFont.setFontName(font);//字体样式：黑体、宋体等
     * hssfFont.setFontHeightInPoints((short)size);//字体大小，short格式
     * hssfFont.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//设置粗体，不需要可不设
     *
     * style.setBorderLeft(HSSFCellStyle.BORDER_THIN);//设置左边框
     * style.setBorderBottom(HSSFCellStyle.BORDER_THIN);//设置下边框
     * style.setBorderRight(HSSFCellStyle.BORDER_THIN);//设置右边框
     * style.setBorderTop(HSSFCellStyle.BORDER_THIN);//设置上边框
     * style.setDataFormat(HSSFDataFormat.getBuiltinFormat("text"));//设置数据格式：文本、货币、日期，百分比等
     * HSSFDataFormat.getBuiltinFormat("(#,##0.00_);[Red](#,##0.00)")//创建数字格式，负数则为红色显示。传入的值必须为double等数字类型
     * HSSFDataFormat.getBuiltinFormat("text")//创建文本格式
     * HSSFDataFormat.getBuiltinFormat("yyyy-m-d")//创建日期格式，该方式展示的数据必须转为String类型，同时制作出来的excel表仍展示为文本类型，需要双击数据格才转为日期格式。建议采用下一种方法。
     * HSSFDataFormat.getBuiltinFormat("0.00%")//创建百分比格式
     * //创建日期格式数据，该方法传入的参数为Date类型，制作excel展示出去即为日期格式。
     * HSSFDataFormat format= wb.createDataFormat();
     * style.setDataFormat(format.getFormat("yyyy/m/d"));//括号中参数为指定的日期类型的格式。
     *
     */
}
