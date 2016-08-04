import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by Heng Song on 8/4/2016.
 */
public class FillExcel {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook book = WorkbookFactory.create(ClassLoader.getSystemResourceAsStream("ngoal.xlsx"));
        InputStream lengthIn = ClassLoader.getSystemResourceAsStream("height.xlsx");

        final InputStream typeStream = ClassLoader.getSystemResourceAsStream("type2.xlsx");
        Map<LengthNode, List<String>> heightMap = loadLength(lengthIn);
        Map<LengthNode, String[]> typeMap = loadType(typeStream);
        lengthIn.close();
        typeStream.close();
        LengthNode node = new LengthNode(20, 44);
        System.out.println(typeMap.containsKey(node));
        System.out.println(heightMap.containsKey(node));
        final List<Integer> heightRows = getHeightRow(book);
        System.out.println(heightRows.contains(2950));
        Collections.sort(heightRows, (o1, o2) -> o2 - o1);
        System.out.println(heightRows.get(0));
        fill(book.getSheetAt(0), heightRows,heightMap,typeMap);

        FileOutputStream fos = new FileOutputStream("FilledLength.xlsx");
        book.write(fos);
        fos.flush();
        fos.close();

    }

    private static void fill(Sheet sheet, List<Integer> rows, Map<LengthNode, List<String>> lengthMap,Map<LengthNode,String[]> typeMap) {

        int startCol = 3;
        for(Integer row : rows ){
            Row type = sheet.getRow(row - 1);
            Row height = sheet.getRow(row);
            Cell cell = height.getCell(1);
            String range = cell.getStringCellValue();
            if(!range.contains("-")){
                Cell tmp = sheet.getRow(row - 1).getCell(1);
                if(tmp == null) System.out.println("**************" + (row - 1) + ":" + 1);
                range = tmp.getStringCellValue();
            }
            String[] tmp = range.split("-");
            String firstPart = tmp[0].trim();
            String secondPart = tmp[1].trim();
            if(firstPart.isEmpty() || secondPart.isEmpty() ){
                System.out.println("row " + row +" is empty ." + range+" :" + firstPart +": " + secondPart);
            }
            LengthNode node = new LengthNode(Integer.parseInt(firstPart), Integer.parseInt(secondPart));
            List<String> heights = lengthMap.get(node);
            String[] arr = typeMap.get(node);
            for (int i = 0; i < heights.size(); i++) {
                Cell nCell = height.getCell(startCol + i);
                Cell typeCell = type.getCell(startCol + i);
                if(nCell == null ) {
                    System.out.println("Row " + row +", Cell " + (startCol+i) +" is null  .   Node" + node);
                    nCell = height.createCell(startCol + i);
                }
                if(getCellValue(nCell).isEmpty()){
                    nCell.setCellValue(heights.get(i));
                }
                if(typeCell == null){
                    typeCell = type.createCell(startCol + i);
                }
                if(getCellValue(typeCell).isEmpty()){
                    typeCell.setCellValue(arr[i]);
                }

            }
        }
    }
    private static String getCellValue(Cell cell){
        switch (cell.getCellType()){
            case Cell.CELL_TYPE_NUMERIC:
                Double val = cell.getNumericCellValue();
                return String.valueOf(val.intValue());
            case Cell.CELL_TYPE_STRING:
                return cell.getStringCellValue();
            default:
                return "";
        }
    }
    private static List<Integer> getHeightRow(Workbook book) {
        assert book != null;
        List<Integer> list = new ArrayList<>();
        Sheet sheet = book.getSheetAt(0);
        for(int i = 0 ; i < sheet.getLastRowNum(); i ++) {
            Row row = sheet.getRow(i);
            if(row == null || row.getPhysicalNumberOfCells()<2) continue;
            Cell cell = sheet.getRow(i).getCell(2);
            String value = cell.getStringCellValue();
            if(value.trim().equalsIgnoreCase("height")){
                list.add(i);
            }
        }
        return list;
    }
    private static  Map<LengthNode,List<String>> loadLength(InputStream in ){
        Map<LengthNode, List<String>> map = new HashMap<>();
        try {
            Workbook book = WorkbookFactory.create(in);
            Sheet sheet = book.getSheetAt(0);
            int row = 0 ;
            while (row < sheet.getLastRowNum()) {
                String name = sheet.getRow(row).getCell(0).getStringCellValue();
                String[] tmp = name.replace("{", "").replace("}", "").split(";");
                LengthNode node = new LengthNode(Integer.parseInt(tmp[0].trim()),Integer.parseInt(tmp[1].trim()));
                List<String> list = new ArrayList<>();
                for(int i = 1 ; i <= 6 ; i ++  ){
                    String cellValue = sheet.getRow(row + i).getCell(0).getStringCellValue();
                    cellValue = cellValue.substring(cellValue.indexOf(".")+1).replace("NA","N/A").trim();
                    if(cellValue.contains(".")) cellValue = cellValue.substring(0, cellValue.indexOf("."));
                    list.add(cellValue);
                }
                map.put(node, list);
                row+=7;
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return map;
    }

    private static Map<LengthNode, String[]> loadType(InputStream in) {
        Map<LengthNode, String[]> map = new HashMap<>();
        try {
            Workbook book = WorkbookFactory.create(in);
            Sheet sheet = book.getSheetAt(0);
            for (int i = 0; i < sheet.getLastRowNum(); i += 2) {
                Row row = sheet.getRow(i);
                assert row != null;
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    if(cell == null) continue;
                    if(cell.getStringCellValue().contains("{")){
                        String range = cell.getStringCellValue();
                        String[] tmp = range.replace("{","").replace("}","").split(";");
                        assert tmp.length == 2;
                        String firstPart = tmp[0].trim(), secondPart = tmp[1].trim();
                        LengthNode node = new LengthNode(Integer.parseInt(firstPart), Integer.parseInt(secondPart));
                        if(!map.containsKey(node)) map.put(node, new String[6]);
                        String type = sheet.getRow(i + 1).getCell(j).getStringCellValue();
                        type = type.trim();
                        String[] arr = type.split("\\s+");
                        String idx = arr[1].substring(0,1);
                        Integer index = Integer.parseInt(idx);
                        map.get(node)[index] = arr[2].trim();
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        } catch (InvalidFormatException e) {
            e.printStackTrace();
        }
        return map;
    }

    static class LengthNode{
        int first;
        int last;

        public LengthNode(int first, int last) {
            this.first = first;
            this.last = last;
        }

        @Override
        public String toString() {
            return "LengthNode{" +
                    "first=" + first +
                    ", last=" + last +
                    '}';
        }

        @Override
        public boolean equals(Object o) {
            if (this == o) return true;
            if (!(o instanceof LengthNode)) return false;

            LengthNode that = (LengthNode) o;

            if (first != that.first) return false;
            return last == that.last;

        }

        @Override
        public int hashCode() {
            int result = first;
            result = 31 * result + last;
            return result;
        }
    }
}
