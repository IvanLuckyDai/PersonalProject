import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;
import java.util.Scanner;

public class Main {

    private static XSSFWorkbook new_workbook = new XSSFWorkbook();
    private static XSSFSheet new_sheet;
    private static int sheet_num;

    public static void main(String[] args) throws IOException, InterruptedException {
//        Welcome();
        initExcel(); //初始化合成工作簿
        File[] fs = getAllFiles("sourceExcel"); //获取所有Excel文件
        List<String> allHeads = getAllHead(fs);
        writeAllHeadToNewExcel(new_workbook, allHeads);
        int i = 1;
        long startTime = System.currentTimeMillis();
        System.out.println("==================开始合并数据==================");
        for (File f : fs) {
            writeDataToNewExcel(f);
            System.out.println("==================已完成" + (i++) + "/" + fs.length + "===================");
        }
        long endTime = System.currentTimeMillis();
        System.out.println("================合并结束耗时" + (endTime - startTime) / 1000 + "秒================");
        generateFile("合并数据.xlsx");

        System.out.println("\n五秒后退出程序~~~");
        Thread.sleep(5000);
        System.exit(0);
    }

    private static void Welcome() {
        int functionType, functionTypeCase;

        System.out.println("请正确输入编号，确定您要执行的功能：");
        System.out.println("1. 合表操作（将一些工作表或工作簿合为一个工作簿的工作表");
        System.out.println("2. 分表操作（将一个工作表或工作簿分为几个工作簿或工作表");

        Scanner input = new Scanner(System.in);
        functionType = input.nextInt();
        if (functionType == 1) {
            System.out.println("请正确输入您当前的合表情况编号：");
            System.out.println("1. 将多个工作簿的第几个工作表合并为一个新工作簿的工作表");
            System.out.println("2. 将多个工作簿的所有工作表合并到一个新工作簿内，并按照<工作簿名称-工作表名称>对工作表进行命名");
            System.out.println("3. 将一个工作簿的几个工作表合并到当前工作簿一个新的工作表或一个新的工作簿的工作表");
            System.out.println("4. 将一个工作簿的所有工作表合并到当前工作簿一个新的工作表或一个新的工作簿的工作表");
        } else if (functionType == 2) {
            System.out.println("请正确输入您当前的分表情况编号：");
            System.out.println("1. 将一个工作簿内的所有工作表拆分为独立的工作簿，并使用工作表名称命名新生成的工作簿");
            System.out.println("2. 将多个工作簿内的所有工作表拆分为独立的工作簿，并按照 <工作簿名称-工作表名称.xlsx> 进行命名");
            System.out.println("3. 将一个工作表内的数据按照某列的筛选项拆分到当前工作簿的多个工作表，并按照<筛选项>对工作表进行命名");
            System.out.println("4. 将一个工作表内的数据按照某列的筛选项拆分为多个新的工作簿，并使用<筛选项.xlsx>对工作簿进行命名");
        } else {
            System.out.print("输入有误，");
            Welcome();
        }
    }

    //初始化合成工作簿
    private static XSSFWorkbook initExcel() throws InterruptedException {
        new_sheet = new_workbook.createSheet("Sheet1");
        System.out.println("==================正在初始化中==================");
//        Thread.sleep(2000);
        System.out.println("==================初始化已完毕==================");
        System.out.print("请输入要合并Excel的第几个sheet：");
        Scanner input = new Scanner(System.in);
        sheet_num = input.nextInt();
        return new_workbook;
    }

    //得到表头
    private static List<String> getAllHead(File[] fileList) throws IOException {
        List<String> heads = new LinkedList<String>();
        Workbook workbook = WorkbookFactory.create(fileList[0]);
        Row firstRow = workbook.getSheetAt(sheet_num).getRow(0);
        for (int i = 0; i < firstRow.getPhysicalNumberOfCells(); i++) {
            heads.add(firstRow.getCell(i).toString());
        }
        return heads;
    }

    //将所有表头写入合成Excel
    private static void writeAllHeadToNewExcel(XSSFWorkbook new_workbook, List<String> heads) {
        XSSFRow new_rows = new_workbook.getSheetAt(0).createRow(0);
        for (int i = 0; i < heads.size(); i++) {
            new_rows.createCell(i).setCellValue(heads.get(i));
            new_rows.getCell(i).setCellType(CellType.STRING);
        }
    }

    //遍历所有工作簿
    private static File[] getAllFiles(String DirectoryName) throws InterruptedException {
        File file = new File(new File("").getAbsolutePath() + "/" + DirectoryName);
        System.out.println(file);
        System.out.println("==================正在加载文件==================");
        for (File listFile : file.listFiles()) {
            System.out.println(listFile);
        }
        System.out.println("==================文件加载完毕==================");
        return file.listFiles();
    }

    //将此文件第二行开始写入目标Excel
    private static void writeDataToNewExcel(File file) throws IOException {
        Workbook source_workbook = WorkbookFactory.create(file);
        Sheet source_sheet = source_workbook.getSheetAt(sheet_num - 1);
        for (int row = 1; row < source_sheet.getLastRowNum() + 1; row++) {
            XSSFRow new_rows = new_sheet.createRow(new_sheet.getLastRowNum() + 1);

            for (int col = 0; col < new_sheet.getRow(0).getPhysicalNumberOfCells(); col++) {
                if (source_sheet.getRow(row).getCell(col) == null || source_sheet.getRow(row).getCell(col).getCellType() == CellType.BLANK) {
                    break;
                } else {
                    switch (source_sheet.getRow(row).getCell(col).getCellType()) {
                        case STRING:
                            new_rows.createCell(col).setCellValue(source_sheet.getRow(row).getCell(col).getStringCellValue());
                            new_rows.getCell(col).setCellType(CellType.STRING);
                            break;
                        case NUMERIC:
                            new_rows.createCell(col).setCellValue(source_sheet.getRow(row).getCell(col).getNumericCellValue());
                            new_rows.getCell(col).setCellType(CellType.NUMERIC);
                            break;
                    }
                }
            }
        }
        source_workbook.cloneSheet(0);
        source_workbook.close();
    }


    //生成文件
    private static void generateFile(String fileName) throws IOException {
        File new_file = new File(new File("").getAbsolutePath() + "/" + fileName);
        FileOutputStream outputStream = new FileOutputStream(new_file);
        System.out.println("==================正在生成文件==================");
        new_workbook.write(outputStream);
        System.out.println("==================文件生成完毕==================");
        System.out.println("==================请查看文件夹==================");
        System.out.println("new_file");
    }

}
