package me.starchier;

import com.google.gson.Gson;
import me.starchier.json.Choose;
import me.starchier.json.Question;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

public class Main {
    public static List<File> xlsxList = new ArrayList<>();
    public static final Gson gson = new Gson();
    public static int groups = 0;
    public static void main(String[] args) {
	// write your code here
        Scanner scanner = new Scanner(System.in);
        System.out.println("本程序可将xlsx表格文件转换成适用于“答题小程序”的json格式文本。");
        System.out.println("请将需要读取的xlsx文件放在与jar的同级目录下。");
        System.out.println("转换后json文件将会原地生成。");
        System.out.println("本程序会自动检测本目录下所有符合条件的xlsx文件并尝试进行读取。");
        System.out.println("请注意：如果需要加入多选题，则需要另起一个新的表格！");
        System.out.println("请注意：xlsx的文件名将作为examid的值！");
        System.out.println("\n开发：Starchier");
        System.out.println("程序使用了以下开源库： Google Gson, Apache poi-ooxml\n");
        System.out.println(" 请按下回车键继续...");
        String input = scanner.nextLine();
        input = null;
        System.out.println("╔ 正在搜索XLSX文件...");
        System.out.println("╠ ");
        File list = new File(System.getProperty("user.dir"));
        for(File file : list.listFiles()) {
            if(file.getName().endsWith(".xlsx")) {
                xlsxList.add(file);
            }
        }
        System.out.println("╠  · 一共找到 " + xlsxList.size() + " 个XLSX文件。");
        System.out.println("╠ ");
        System.out.println("╠ 开始处理XLSX文件...");
        for(File file : xlsxList) {
            System.out.println("╠ ");
            List<Question> questions = new ArrayList<>();
            System.out.println("╠ ╔ 正在读取文件 " + file.getName());
            XSSFWorkbook xlsFile = null;
            try {
                xlsFile = new XSSFWorkbook(new FileInputStream(file));
            } catch (IOException e) {
                e.printStackTrace();
            }
            XSSFSheet sheet = xlsFile.getSheetAt(0);
            for(Row row : sheet) {
                if(row.getRowNum() == 0) {
                    processHeader(row);
                    System.out.println("╠ ╠  · 该文件中每道题一共有 " + groups + "组选项。");
                    continue;
                }
                String[] element = new String[5];
                String[] chooses = new String[groups * 3];
                int i = 0;
                for(Cell cell : row) {
                    if(cell.getColumnIndex() < 5) {
                        fitType(row, element, i, cell);

                    } else {
                        fitType(row, chooses, i - 5, cell);
                    }
                    i++;
                }
                Choose[] chooseInstances = new Choose[groups];
                for(int j = 0; j < groups; j++) {
                    if(j == 0) {
                        chooseInstances[j] = new Choose(chooses[0], chooses[1], chooses[2]);
                    } else {
                        chooseInstances[j] = new Choose(chooses[j * 3], chooses[j * 3 + 1], chooses[j * 3 + 2]);
                    }
                }
                questions.add(new Question(element[0], element[1], element[2], element[3], element[4], chooseInstances, file.getName().replace(".xlsx", "")));
            }
            System.out.println("╠ ╠ 正在将文件 " + file.getName() + " 转换为JSON...");
            StringBuilder sb = new StringBuilder();
            for(Question q : questions) {
                sb.append(gson.toJson(q));
                sb.append("\n");
            }
            File jsonFile = new File(file.getName().replace(".xlsx", "") + ".json");
            if(jsonFile.exists()) {
                if(!jsonFile.delete()) {
                    System.out.println("╠ ╠   [警告] 无法删除已存在的json文件 " + jsonFile.getName() + " 这可能会发生未知的特性（bug）!");
                }
            }
            try {
                jsonFile.createNewFile();
                BufferedWriter out = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(jsonFile.getName()), StandardCharsets.UTF_8));
                out.write(sb.toString());
                out.close();
                System.out.println("╠ ╚ 已输出json文件： " + jsonFile.getName());
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        System.out.println("╠ ");
        System.out.println("╚ 程序已处理完毕所有 XLSX 文件。");
    }

    private static void fitType(Row row, String[] chooses, int i, Cell cell) {
        if(cell.getCellType().equals(CellType.STRING)) {
            chooses[i] = cell.getStringCellValue();
        } else if(cell.getCellType().equals(CellType.NUMERIC)) {
            chooses[i] = String.valueOf(cell.getNumericCellValue()).replace(".0", "");
        } else {
            chooses[i] = cell.toString();
            System.out.println("╠ ╠  [警告] 行" + row.getRowNum() + " 列 " + cell.getColumnIndex() + "可能存在问题，请检查。");
        }
    }

    public static void processHeader(Row row) {
        int length = row.getLastCellNum();
        groups = (length - 5) / 3;
    }
}
