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

    static XSSFWorkbook new_workbook = new XSSFWorkbook();
    static XSSFSheet new_sheet;
    static int sheet_num;

    public static void main(String[] args) throws IOException, InterruptedException {
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

    //初始化合成工作簿
    static XSSFWorkbook initExcel() throws InterruptedException {
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
    static List<String> getAllHead(File[] fileList) throws IOException {
        List<String> heads = new LinkedList<String>();
        Workbook workbook = WorkbookFactory.create(fileList[0]);
        Row firstRow = workbook.getSheetAt(0).getRow(0);
        for (int i = 0; i < firstRow.getPhysicalNumberOfCells(); i++) {
            heads.add(firstRow.getCell(i).toString());
        }
        return heads;
    }

    //将所有表头写入合成Excel
    static void writeAllHeadToNewExcel(XSSFWorkbook new_workbook, List<String> heads) {
        XSSFRow new_rows = new_workbook.getSheetAt(0).createRow(0);
        for (int i = 0; i < heads.size(); i++) {
            new_rows.createCell(i).setCellValue(heads.get(i));
            new_rows.getCell(i).setCellType(CellType.STRING);
        }
    }

    //遍历所有工作簿
    static File[] getAllFiles(String DirectoryName) throws InterruptedException {
        File file = new File(new File("").getAbsolutePath() + "\\" + DirectoryName);
        System.out.println("==================正在加载文件==================");
        for (File listFile : file.listFiles()) {
            System.out.println(listFile);
//            Thread.sleep(1000);
        }
        System.out.println("==================文件加载完毕==================");
        return file.listFiles();
    }

    //将此文件第二行开始写入目标Excel
    static void writeDataToNewExcel(File file) throws IOException {
        Workbook source_workbook = WorkbookFactory.create(file);
        Sheet source_sheet = source_workbook.getSheetAt(sheet_num - 1);
//        System.out.println(file + "=========source" + source_sheet.getLastRowNum());
        for (int row = 1; row < source_sheet.getLastRowNum() + 1; row++) {
            XSSFRow new_rows = new_sheet.createRow(new_sheet.getLastRowNum() + 1);

//            System.out.println("总表目前一共" + (new_sheet.getLastRowNum() + 1) + "行");
            for (int col = 0; col < new_sheet.getRow(0).getPhysicalNumberOfCells(); col++) {
//                System.out.println(new_sheet.getRow(0).getPhysicalNumberOfCells() + "     col" + col + " : " + new_sheet.getRow(0).getCell(col).toString());
//                System.out.println(source_sheet.getRow(row).getCell(col).getCellType());
                if (source_sheet.getRow(row).getCell(col) == null || source_sheet.getRow(row).getCell(col).getCellType() == CellType.BLANK) {
//                    System.out.println("NULL------------------");
                    break;
                } else {
                    switch (source_sheet.getRow(row).getCell(col).getCellType()) {
                        case STRING:
                            new_rows.createCell(col).setCellValue(source_sheet.getRow(row).getCell(col).getStringCellValue());
                            new_rows.getCell(col).setCellType(CellType.STRING);
//                            System.out.print(source_sheet.getRow(row).getCell(col).getStringCellValue() + "\t");
                            break;
                        case NUMERIC:
                            new_rows.createCell(col).setCellValue(source_sheet.getRow(row).getCell(col).getNumericCellValue());
                            new_rows.getCell(col).setCellType(CellType.NUMERIC);
//                            System.out.print(source_sheet.getRow(row).getCell(col).getNumericCellValue() + "\t");
                            break;
                    }
                }
            }
//            System.out.println();
        }
        source_workbook.cloneSheet(0);
        source_workbook.close();
    }


    //生成文件
    static void generateFile(String fileName) throws IOException {
        File new_file = new File(new File("").getAbsolutePath() + "\\" + fileName);
        FileOutputStream outputStream = new FileOutputStream(new_file);
        System.out.println("==================正在生成文件==================");
        new_workbook.write(outputStream);
        System.out.println("==================文件生成完毕==================");
        System.out.println("==================请查看文件夹==================");
        System.out.println("new_file");
    }

}
