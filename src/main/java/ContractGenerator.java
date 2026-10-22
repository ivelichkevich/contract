import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.nio.file.*;
import java.util.*;

public class ContractGenerator {

    public static void main(String[] args) {
        if (args.length == 0) {
            System.err.println("Использование: java ContractGenerator <путь_к_файлу.properties>");
            System.exit(1);
        }

        String propsPath = args[0];
        File propsFile = new File(propsPath);
        if (!propsFile.exists()) {
            System.err.println("Ошибка: Файл не найден по пути: " + propsPath);
            System.exit(1);
        }

        // Базовая директория, где лежит файл свойств
        Path baseDir = propsFile.getAbsoluteFile().getParentFile().toPath();
        // Папка с шаблонами Word
        Path templatesDir = Paths.get("templates"); 

        try {
            Properties props = new Properties();
            try (InputStream input = new FileInputStream(propsFile)) {
                props.load(input);
            }
            
            // ХИТРОСТЬ: Если в properties есть income, но нет salary, приравниваем их
            if (props.getProperty("salary") == null && props.getProperty("income") != null) {
                props.setProperty("salary", props.getProperty("income"));
            }
            
            String firstName = props.getProperty("first-name", "Unknown");
            String lastName = props.getProperty("last-name", "Unknown");
            
            // Поиск подписи: папка Имя_Фамилия/sign.png
            String folderName = firstName + "_" + lastName;
            Path signaturePath = baseDir.resolve(folderName).resolve("sign.png");
            File signatureFile = signaturePath.toFile();
            
            if (!signatureFile.exists()) {
                System.out.println("ВНИМАНИЕ: Подпись не найдена: " + signaturePath.toAbsolutePath());
            } else {
                System.out.println("Найдена подпись: " + signaturePath.toAbsolutePath());
            }

            // Проверяем наличие папки с шаблонами
            File[] templateFiles = templatesDir.toFile().listFiles((dir, name) -> name.endsWith(".docx"));
            if (templateFiles == null || templateFiles.length == 0) {
                System.err.println("Ошибка: В папке 'templates' нет .docx файлов!");
                System.exit(1);
            }

            // Проходимся по каждому шаблону в папке
            for (File template : templateFiles) {
                String docName = "Generated_" + template.getName().replace(".docx", "") + "_" + lastName + ".docx";
                Path outputPath = baseDir.resolve(docName);
                
                generateFromTemplate(template, outputPath.toFile(), props, signatureFile);
                System.out.println("Сгенерирован: " + outputPath.toAbsolutePath());
            }
            
            System.out.println("Готово! Все документы созданы.");
            
        } catch (Exception e) {
            System.err.println("Произошла ошибка при генерации:");
            e.printStackTrace();
        }
    }

    private static void generateFromTemplate(File templateFile, File outputFile, Properties props, File signatureFile) throws Exception {
        try (XWPFDocument document = new XWPFDocument(new FileInputStream(templateFile))) {
            
            // 1. Обработка обычных параграфов
            for (XWPFParagraph p : document.getParagraphs()) {
                replaceTextInParagraph(p, props, signatureFile);
            }

            // 2. Обработка таблиц
            for (XWPFTable table : document.getTables()) {
                for (XWPFTableRow row : table.getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph p : cell.getParagraphs()) {
                            replaceTextInParagraph(p, props, signatureFile);
                        }
                    }
                }
            }

            // Сохранение
            try (FileOutputStream out = new FileOutputStream(outputFile)) {
                document.write(out);
            }
        }
    }

    private static void replaceTextInParagraph(XWPFParagraph p, Properties props, File signatureFile) throws Exception {
        String paragraphText = p.getText();
        if (paragraphText == null || paragraphText.isEmpty()) return;

        boolean hasReplacement = false;

        // 1. Обработка текстовых плейсхолдеров
        if (paragraphText.contains("${")) {
            for (String key : props.stringPropertyNames()) {
                String value = props.getProperty(key);
                
                // Ищем точное совпадение: ${first-name}
                String placeholder1 = "${" + key + "}";
                if (paragraphText.contains(placeholder1)) {
                    paragraphText = paragraphText.replace(placeholder1, value);
                    hasReplacement = true;
                }
                
                // Ищем с подчеркиванием (на всякий случай): ${first_name}
                String placeholder2 = "${" + key.replace("-", "_") + "}";
                if (paragraphText.contains(placeholder2)) {
                    paragraphText = paragraphText.replace(placeholder2, value);
                    hasReplacement = true;
                }
            }
        }

        // 2. Обработка подписи
        boolean hasSignature = paragraphText.contains("[SIGNATURE]");
        if (hasSignature) {
            paragraphText = paragraphText.replace("[SIGNATURE]", "");
            hasReplacement = true;
        }

        // 3. Если были замены, перезаписываем абзац
        if (hasReplacement && !p.getRuns().isEmpty()) {
            // Запоминаем форматирование (жирность, шрифт) первого куска текста
            XWPFRun firstRun = p.getRuns().get(0);
            boolean isBold = firstRun.isBold();
            boolean isItalic = firstRun.isItalic();
            String fontFamily = firstRun.getFontFamily();
            int fontSize = firstRun.getFontSize();

            // Полностью очищаем абзац от старых "рваных" кусков (Word Run Splitting)
            for (int i = p.getRuns().size() - 1; i >= 0; i--) {
                p.removeRun(i);
            }

            // Вставляем очищенный и замененный текст
            if (!paragraphText.trim().isEmpty()) {
                XWPFRun newRun = p.createRun();
                newRun.setText(paragraphText);
                newRun.setBold(isBold);
                newRun.setItalic(isItalic);
                if (fontFamily != null) newRun.setFontFamily(fontFamily);
                if (fontSize != -1) newRun.setFontSize(fontSize);
            }

            // Если была метка подписи, вставляем картинку
            if (hasSignature) {
                XWPFRun sigRun = p.createRun();
                if (signatureFile.exists()) {
                    try (InputStream picStream = new FileInputStream(signatureFile)) {
                        sigRun.addPicture(picStream, Document.PICTURE_TYPE_PNG, signatureFile.getName(),
                                Units.toEMU(150), Units.toEMU(50));
                    }
                } else {
                    sigRun.setText("_________________");
                }
            }
        }
    }
}