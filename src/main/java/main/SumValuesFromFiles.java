package main;

import com.spire.xls.CellRange;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import java.io.File;
import java.util.HashMap;


public class SumValuesFromFiles {
    public static void main(String[] args) {
//        Scanner sc = new Scanner(System.in);
//        System.out.println("Path para a pasta com os ficheiros de estimativas?");
//        String folderPath = sc.nextLine();
        String folderPath = "C:\\Users\\dmonteim\\NTT DATA EMEAL\\[CGD]_AM_Canais_Ano_1_(AIS)_(005002) - Testes automáticos\\sandbox_testing"; // Change the folder path as per your requirement

        String sheetName = "Resumo"; // Change the sheet name as per your requirement
        String cellReference = "D26"; // Change the cell reference as per your requirement

        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        HashMap<String, Double> output = new HashMap<>();

        assert files != null;
        for (File file : files) {
            String fileName = file.getName();
            double currentSum = getCellValue(String.valueOf(file), sheetName, cellReference);
            if (file.getName().contains("__")){
                if (!output.containsKey("Sem epic associado")){
                    output.put("Sem Épico associado", currentSum);
                }else {
                    output.put("Sem Épico associado", output.get("Sem Épico associado") + currentSum);
                }
            }else {
                String epic_name = fileName.substring(0,fileName.indexOf("_"))+"_"+fileName.substring(fileName.indexOf("_")+1,fileName.indexOf("_",fileName.indexOf("_")+1));
                if (!output.containsKey(epic_name)){
                    output.put(epic_name, currentSum);
                }else{
                    output.put(epic_name, output.get(epic_name)+currentSum);
                }
            }
            System.out.println(output);
        }

        System.out.println(output);
        }

    public static double getCellValue(String fileName, String sheetName, String Range){
        String password = "NTT DATA"; // Change the sheet password as per your requirement

        Workbook workbook = new Workbook();
        workbook.loadFromFile(fileName);

        Worksheet sheet = workbook.getWorksheets().get(sheetName);

        sheet.unprotect(password);

        CellRange cell = sheet.getRange().get(Range);

        sheet.protect(password);

        return (double) cell.getFormulaValue();
    }

}
