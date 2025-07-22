import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import org.apache.poi.ss.usermodel.DateUtil;

public class ExcelToTxtConverter {
    private static final DateTimeFormatter outputFormat = DateTimeFormatter.ofPattern("dd/MM/yy HH:mm");

    public static void main(String[] args) {
        String excelFilePath = "C://piratas//Quadro_de_Provas.xlsx";
        String txtFilePath = "C://piratas//teste.txt";
        try (
             FileInputStream fis = new FileInputStream(excelFilePath);
             Workbook workbook = new XSSFWorkbook(fis);
             BufferedWriter writer = new BufferedWriter(new FileWriter(txtFilePath))) {

            Sheet sheet = workbook.getSheetAt(0); // Lê a primeira planilha
            StringBuilder builder = new StringBuilder();
            builder.append("Entrega       | Tipo                 | Prova                                               | Tempo restante\n");
            builder.append("-------------------------------------------------------------------------------------------------------------\n");
            LocalDateTime agora = LocalDateTime.now();
            String tempoRestante;
            for (Row row : sheet) {
                if(row.getRowNum() > 0) {
                    LocalDateTime dataHora = row.getCell(2).getLocalDateTimeCellValue();
                    Duration duracao = Duration.between(agora,dataHora);

                    if (!duracao.isNegative()) {
                        long dias = duracao.toDays();
                        long horas = duracao.minusDays(dias).toHours();
                        long minutos = duracao.minusDays(dias).minusHours(horas).toMinutes();
                        tempoRestante = String.format("%dd %02dh %02dm", dias, horas, minutos);
                    } else {
                        tempoRestante = "⏱ Finalizado";
                    }
                    builder.append(String.format("%s | %-20s | %-50s | %s\n",
                            dataHora.format(outputFormat), row.getCell(0).getStringCellValue(), row.getCell(1).getStringCellValue(), tempoRestante));
                }
            }
            writer.write(builder.toString());
            writer.flush();
            System.out.println("Arquivo TXT gerado com sucesso!");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static String getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return cell.getDateCellValue().toString();
                } else {
                    return String.valueOf(cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            case BLANK:
                return "";
            default:
                return "";
        }
    }
}
