import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.nio.file.*;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import static java.nio.file.StandardWatchEventKinds.*;

public class ExcelToTxtConverter {
    private static final DateTimeFormatter outputFormat = DateTimeFormatter.ofPattern("dd/MM/yy HH:mm:ss");


    private static final String EXCEL_FILE = "C://piratas//Quadro_de_Provas.xlsx";
    private static final String TXT_FILE = "teste.txt";

    private static final Path EXCEL_PATH = Paths.get("C://piratas//Quadro_de_Provas.xlsx");
    private static final Path TXT_PATH = Paths.get("teste.txt");
    private static final DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public static void main(String[] args) throws IOException, InterruptedException {
        ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
        Runnable task = () -> {
            convertExcelToTxt(EXCEL_FILE, TXT_FILE);
        };

        scheduler.scheduleAtFixedRate(task, 0, 1, TimeUnit.SECONDS);

        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            scheduler.shutdown();
            System.out.println("Scheduler encerrado.");
        }));
    }

    public static void convertExcelToTxt(String excelPath, String txtPath) {
        try (
             FileInputStream fis = new FileInputStream(excelPath);
             Workbook workbook = new XSSFWorkbook(fis);
             BufferedWriter writer = new BufferedWriter(new FileWriter(txtPath))) {

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
                        long seconds = duracao.getSeconds();
                        tempoRestante = String.format("%dd %02dh %02dm %02ds", dias, horas, minutos, (seconds / 1000) % 60);
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
}
