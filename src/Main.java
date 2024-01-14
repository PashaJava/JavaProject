import java.io.*;
import java.util.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.AxisCrosses;
import org.apache.poi.xddf.usermodel.chart.AxisPosition;
import org.apache.poi.xddf.usermodel.chart.ChartTypes;
import org.apache.poi.xddf.usermodel.chart.LegendPosition;
import org.apache.poi.xddf.usermodel.chart.MarkerStyle;
import org.apache.poi.xddf.usermodel.chart.XDDFCategoryAxis;
import org.apache.poi.xddf.usermodel.chart.XDDFChartLegend;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFDataSourcesFactory;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFNumericalDataSource;
import org.apache.poi.xddf.usermodel.chart.XDDFValueAxis;
import org.apache.poi.xssf.usermodel.XSSFChart;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

    // Создание списков с отличниками, хорошистами, троечниками и не допущенными
    private static ArrayList<String> Otl = new ArrayList<>();
    private static ArrayList<String> Hor = new ArrayList<>();
    private static ArrayList<String> Udovl = new ArrayList<>();
    private static ArrayList<String> Neud = new ArrayList<>();

    // Создание счетчиков кол-ва отличников, хорошистов, троечников и не допущенных
    private static double Five = 0;
    private static double Four = 0;
    private static double Three = 0;
    private static double Two = 0;


    public static void main(String[] args) throws Exception {
        try {

            // Чтение данных из исходного Excel файла по строкам
            HashMap<String, Double> students = new HashMap<>();
            String path = args[0];
            FileInputStream file = new FileInputStream(new File(path));
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            XSSFSheet sheet = workbook.getSheetAt(0);
            Iterator<Row> rowIterator = sheet.iterator();
            while (rowIterator.hasNext()) {
                String student = "";
                double mark = 0;
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                String guy = "";
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()) {  // Проверка читаемой ячейки на тип переменной
                        case NUMERIC:
                            double num = cell.getNumericCellValue();
                            mark = num;
                            if (num == 5) {
                                Five++;
                                Otl.add(guy);
                            } else if (num == 4) {
                                Four++;
                                Hor.add(guy);
                            } else if (num == 3) {
                                Three++;
                                Udovl.add(guy);
                            } else if (num == 2) {
                                Two++;
                                Neud.add(guy);
                            // Проверка правильности оценки
                            } else if (num > 5 || num < 2) {
                                System.out.println(" Ошибка в оценке");
                                System.exit(0);
                            }
                            break;
                        case STRING:
                            guy = cell.getStringCellValue();
                            student = guy;
                            break;
                    }
                    students.put(student, mark);
                }
            }

            // Подсчет среднего балла вызовом ф-ции middle
            double mid = middle(Five, Four, Three, Two);
            file.close();
            writeIntoExcel(Five, Four, Three, Two, mid);
        }
        catch (Exception e){
            System.out.println(e.getMessage());
        }

    }

    // Метод для подсчета среднего балла
    public static double middle(double Five, double Four, double Three, double Two) throws Exception {
        try {
            double mid = 0;
            mid = (Five * 5 + Four * 4 + Three * 3 + Two * 2) / (Five + Four + Three + Two);
            return mid;
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
            return 0;
        }
    }

    // Метод для создания новой Excel таблицы для вывода обработанных данных и построение статистического графика
    public static void writeIntoExcel(double Five, double Four, double Three, double Two, double mid) throws Exception {
        try {
            XSSFWorkbook book = new XSSFWorkbook();
            XSSFSheet sheet = book.createSheet("ИБС-22");


            Row row = sheet.createRow(0);
            Cell name = row.createCell(0);
            name.setCellValue("Кол-во отличников:");
            Cell five = row.createCell(1);
            five.setCellValue(Five);

            Row row1 = sheet.createRow(1);
            Cell name1 = row1.createCell(0);
            name1.setCellValue("Кол-во хорошистов:");
            Cell four = row1.createCell(1);
            four.setCellValue(Four);

            Row row2 = sheet.createRow(2);
            Cell name2 = row2.createCell(0);
            name2.setCellValue("Кол-во троечников:");
            Cell three = row2.createCell(1);
            three.setCellValue(Three);

            Row row3 = sheet.createRow(3);
            Cell name3 = row3.createCell(0);
            name3.setCellValue("Кол-во недопущенных:");
            Cell two = row3.createCell(1);
            two.setCellValue(Two);

            Row row4 = sheet.createRow(4);
            Cell name4 = row4.createCell(0);
            name4.setCellValue("Средний балл группы:");
            Cell middle = row4.createCell(1);
            middle.setCellValue(mid);


            Row row6 = sheet.createRow(6);
            Cell otl = row6.createCell(0);
            otl.setCellValue(5);

            Cell hor = row6.createCell(1);
            hor.setCellValue(4);

            Cell udovl = row6.createCell(2);
            udovl.setCellValue(3);

            Cell neud = row6.createCell(3);
            neud.setCellValue(2);

            int max = Math.max(Otl.size(), Math.max(Hor.size(), Math.max(Udovl.size(), Neud.size())));
            int stroka = 7;

            // Построчный вывод фамилий студентов в зависимости от полученной оценки (столбика)
            for (int i = 0; i < max; i++) {
                Row row52 = sheet.createRow(stroka);
                for (int j = 0; j < 4; j++) {

                    if (j == 0 && Otl.size() > i) {
                        Cell cell = row52.createCell(j);
                        cell.setCellValue(Otl.get(i));
                    } else if (j == 1 && Hor.size() > i) {
                        Cell cell = row52.createCell(j);
                        cell.setCellValue(Hor.get(i));
                    } else if (j == 2 && Udovl.size() > i) {
                        Cell cell = row52.createCell(j);
                        cell.setCellValue(Udovl.get(i));
                    } else if (j == 3 && Neud.size() > i) {
                        Cell cell = row52.createCell(j);
                        cell.setCellValue(Neud.get(i));
                    }
                }
                stroka++;
            }


            // Построение статистического графика
            XSSFDrawing drawing = sheet.createDrawingPatriarch();
            // Назначение размера графика и ключевых ячеек(начало графика и конец)
            XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 0, 20, 17, 40);
            // Создание Легенд для графика
            XSSFChart chart = drawing.createChart(anchor);
            XDDFChartLegend legend = chart.getOrAddLegend();
            legend.setPosition(LegendPosition.TOP_RIGHT);
            // Создание осей X и Y и названий осей
            XDDFCategoryAxis bottomAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            bottomAxis.setTitle("Оценка");
            XDDFValueAxis leftAxis = chart.createValueAxis(AxisPosition.LEFT);
            leftAxis.setTitle("Кол-во студентов");
            leftAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // Выбор ячеек для забора данных для осей X и Y
            XDDFDataSource<Double> xs = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(6, 6, 0, 3));
            XDDFNumericalDataSource<Double> ys = XDDFDataSourcesFactory.fromNumericCellRange(sheet, new CellRangeAddress(0, 3, 1, 1));

            XDDFLineChartData data = (XDDFLineChartData) chart.createData(ChartTypes.LINE, bottomAxis, leftAxis);
            XDDFLineChartData.Series series1 = (XDDFLineChartData.Series) data.addSeries(xs, ys);
            series1.setTitle("Успеваемость", null);
            series1.setSmooth(false);
            series1.setMarkerStyle(MarkerStyle.STAR);
            chart.plot(data);


            book.write(new FileOutputStream("Output.xlsx"));
            book.close();
        }
        catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }
}