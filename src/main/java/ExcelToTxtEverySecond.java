import java.io.*;
        import java.nio.file.*;
        import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.concurrent.*;
        import org.apache.poi.ss.usermodel.*;
        import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelToTxtEverySecond {
    private static final Path EXCEL_PATH = Paths.get("dados.xlsx");
    private static final Path TXT_PATH = Paths.get("saida.txt");
    private static final DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public static void main(String[] args) {
        ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
        Runnable task = () -> {
            try (InputStream inp = Files.newInputStream(EXCEL_PATH);
                 Workbook wb = new XSSFWorkbook(inp);
                 BufferedWriter writer = Files.newBufferedWriter(TXT_PATH,
                         StandardOpenOption.CREATE,
                         StandardOpenOption.TRUNCATE_EXISTING)) {


                writer.write("Gerado em: " + LocalDateTime.now().format(dtf));
                writer.newLine();
                System.out.println("Arquivo TXT atualizado: " + LocalDateTime.now().format(dtf));
            } catch (IOException e) {
                e.printStackTrace();
            }
        };

        scheduler.scheduleAtFixedRate(task, 0, 1, TimeUnit.SECONDS);

        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            scheduler.shutdown();
            System.out.println("Scheduler encerrado.");
        }));
    }
}
