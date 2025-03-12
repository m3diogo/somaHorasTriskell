package main;

import com.spire.xls.CellRange;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import me.tongfei.progressbar.ProgressBar;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.HashMap;
import java.util.Scanner;


public class SumValuesFromFiles {
    public static void main(String[] args) {
        Scanner sc = new Scanner(System.in, StandardCharsets.UTF_8);
        System.out.println("""
                #####################################################################################\s
                Esta ferramenta assume que os ficheiros das estimativas estão no formato recomendado:\s
                    Ex: TCAUT_1978_TCAUT_1447_17.11.2023.xlsx\s
                    Onde:\s
                        1. TCAUT_1978 é o Épico\s
                        2. TCAUT_1447 é o Issue\s
                        3. 17.11.2023 é a data de criação (dd.MM.aaaa)\s
                #####################################################################################\s
                \s""");

        System.out.println("Qual é o path absoluto para a pasta com os ficheiros de estimativas? \n" +
                "(Nota: o path não pode conter caracteres especiais, ex. acentos)");

        String folderPath = sc.nextLine();
        System.out.println("Qual é a referência da célula do Esforço Total em Horas?");
        String cellReference = sc.nextLine();
//        String folderPath = "C:\\Users\\dmonteim\\NTT DATA EMEAL\\[CGD]_AM_Canais_Ano_1_(AIS)_(005002) - Testes automáticos\\sandbox_testing"; // Change the folder path as per your requirement

        String sheetName = "Resumo"; // Change the sheet name as per your requirement
//        String cellReference = "D26"; // Change the cell reference as per your requirement
        int numeroFicheiros;

        File folder = new File(folderPath);
        File[] files = folder.listFiles();

        HashMap<String, Double> outputHoras = new HashMap<>();
        try {
            assert files != null;
            numeroFicheiros = files.length;

            System.out.println("Vou iniciar a contagem...");
            try(ProgressBar pb = new ProgressBar("A somar horas:", numeroFicheiros)) {
                for (File file : files) {
                    String fileName = file.getName();
                    double currentSum = getCellValue(String.valueOf(file), sheetName, cellReference);
                    if (file.getName().contains("__")) {
                        if (!outputHoras.containsKey("Sem epic associado")) {
                            outputHoras.put("Sem Épico associado", currentSum);
                        } else {
                            outputHoras.put("Sem Épico associado", outputHoras.get("Sem Épico associado") + currentSum);
                        }
                    } else {
                        String epic_name = fileName.substring(0, fileName.indexOf("_")) + "_" + fileName.substring(fileName.indexOf("_") + 1, fileName.indexOf("_", fileName.indexOf("_") + 1));
//                us_name = fileName.substring(11,21);
                        if (!outputHoras.containsKey(epic_name)) {
                            outputHoras.put(epic_name, currentSum);
                        } else {
                            outputHoras.put(epic_name, outputHoras.get(epic_name) + currentSum);
                        }
                    }
                    pb.step();
                    pb.maxHint(numeroFicheiros);
                }
            }


            for (HashMap.Entry<String, Double> entry : outputHoras.entrySet()){
                if (entry.getKey().equals("Sem Épico associado")){
                    printToFile(entry.getKey() + ": " + entry.getValue() + "h");
                }else {
                    printToFile("Épico " + entry.getKey() + ": " + entry.getValue() + "h");
                }
//                System.out.println("Épico "+ entry.getKey() + ": " + entry.getValue() + "h");
            }

        }catch (NullPointerException e){
            printToFile("""
                    XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    Verifica o path da pasta das estimativas, não pode conter acentos.
                    XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
                    """);

        }
    }

    public static double getCellValue(String fileName, String sheetName, String Range){
        String password = "ntt"; // Change the sheet password as per your requirement

        Workbook workbook = new Workbook();
        workbook.loadFromFile(fileName);

        Worksheet sheet = workbook.getWorksheets().get(sheetName);

        sheet.unprotect(password);

        CellRange cell = sheet.getRange().get(Range);

        sheet.protect(password);

        return (double) cell.getFormulaValue();
    }

    public static void printToFile(String log){
        String fileName = "output.txt";
        try{
            FileWriter writer = new FileWriter(fileName, true);

            writer.append(log);
            writer.append(System.lineSeparator());

            writer.close();
        }catch (IOException e){
            System.err.println("Error: Unable to append to the file.");
        }
    }

}
