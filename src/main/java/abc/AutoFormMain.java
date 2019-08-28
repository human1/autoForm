package abc;

import lombok.extern.log4j.Log4j;
import abc.model.AsiaDJ;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;

@Log4j
public class AutoFormMain extends JFrame implements ActionListener {

    private JPanel panel;
    private JButton btnProcess;
    private JButton btnBrowseFileExcel;
    private JTextArea txaLog;
    private JTextField txtFile;
    private JTextField txtChromeDriver;
    private JButton btnBrowseChrome;
    private JScrollPane scrTxa;

    // Process Excel
    private String path = "";
    private String pathDriver = "";

    // Selenium
    WebDriver driver;
    private static String URL_TO_FORM = "https://docs.google.com/forms/d/e/1FAIpQLScfi7qQQac0JvqhMzNLKXaaPtuh9E32Dj_4tA5-IvKZpB9byg/viewform";

    public static void main(String[] args) {
        try {
            new AutoFormMain();
        } catch (Exception e) {
            log.error(e);
        }
    }

    public AutoFormMain() {
        initFrame();
    }

    private void initFrame() {

        panel = new JPanel();
        panel.setBounds(12, 0, 445, 43);
        getContentPane().add(panel);
        panel.setLayout(null);

        JLabel lblDataFile = new JLabel("Data File");
        lblDataFile.setBounds(10, 81, 70, 13);
        panel.add(lblDataFile);

        txtFile = new JTextField();
        txtFile.setColumns(10);
        txtFile.setBounds(73, 75, 254, 20);
        panel.add(txtFile);

        btnBrowseFileExcel = new JButton("Browse");
        btnBrowseFileExcel.setBounds(349, 74, 84, 21);
        btnBrowseFileExcel.addActionListener(this);
        panel.add(btnBrowseFileExcel);

        txtChromeDriver = new JTextField();
        txtChromeDriver.setColumns(10);
        txtChromeDriver.setBounds(73, 105, 254, 20);
        panel.add(txtChromeDriver);


        JLabel lblChromeDriver = new JLabel("Chrome");
        lblChromeDriver.setBounds(10, 111, 70, 13);
        panel.add(lblChromeDriver);

        btnBrowseChrome = new JButton("Browse");
        btnBrowseChrome.setBounds(349, 105, 84, 21);
        btnBrowseChrome.addActionListener(this);
        panel.add(btnBrowseChrome);

        txaLog = new JTextArea();
        txaLog.setBounds(12, 253, 445, 130);
        panel.add(txaLog);

        btnProcess = new JButton("Process");
        btnProcess.setBounds(73, 155, 84, 41);
        btnProcess.addActionListener(this);
        panel.add(btnProcess);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        pack();
        setSize(470, 470);
        setResizable(true);
        setLocationRelativeTo(null);
        setVisible(true);
        setContentPane(panel);
    }

    @Override
    public void actionPerformed(ActionEvent e) {
        Object src = e.getSource();
        if (src != null) {
            if (src == btnProcess) {
                handleProcess();
            } else if (src == btnBrowseFileExcel) {
                // Open file dialog
                handleOpenFileExcel();
            } else if (src == btnBrowseChrome) {
                handleOpenDriver();
            }
        }
    }

    private void handleProcess() {
        if ("".equals(pathDriver)) {
            JOptionPane.showMessageDialog(panel, "Please select Chrome driver!");
            return;
        }
        if ("".equals(path)) {
            JOptionPane.showMessageDialog(panel, "Please select a file first!");
            return;
        }
        System.setProperty("webdriver.chrome.driver", pathDriver);
        driver = new ChromeDriver();
        Thread r = new ReadExcel(path);
        r.start();
    }

    private void handleOpenFileExcel() {
        JFileChooser jfc = new JFileChooser();
        jfc.setDialogTitle("Select an excel");
        jfc.setAcceptAllFileFilterUsed(false);
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Excel format", "xlsx", "xls");
        jfc.addChoosableFileFilter(filter);
        // Process close dialog
        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            path = jfc.getSelectedFile().getPath();
            txtFile.setText(path);
        }
    }

    private void handleOpenDriver() {
        JFileChooser jfc = new JFileChooser();
        jfc.setDialogTitle("Select Chrome Driver");
        jfc.setAcceptAllFileFilterUsed(false);
        // Process close dialog
        int returnValue = jfc.showOpenDialog(null);
        if (returnValue == JFileChooser.APPROVE_OPTION) {
            pathDriver = jfc.getSelectedFile().getPath();
            txtChromeDriver.setText(pathDriver);
        }
    }

    /**
     * PROCESS WITH EXCEL
     */
    class ReadExcel extends Thread {

        private String pathToFile;

        public ReadExcel(Object parameter) {
            this.pathToFile = (String) parameter;
        }

        @Override
        public void run() {
            log("Start ReadExcel");
            processReadExcelFile(pathToFile);
        }


        private void processReadExcelFile(String path) {
            String ext = path.substring(path.indexOf("."));
            if (".xlsx".equals(ext)) {
                readExcelTypeXLSX(path);
            } else {
                readExcelTypeXLS(path);
            }
        }

        /**
         * Excel 2007+
         *
         * @param pathToFile
         */
        public void readExcelTypeXLSX(String pathToFile) {
            try {
                FileInputStream file = new FileInputStream(new File(pathToFile));
                XSSFWorkbook workbook = new XSSFWorkbook(file);
                XSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) {
                    Row row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    int colIdx = 0;

                    AsiaDJ entity = new AsiaDJ();
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        //
                        processAddCellValue(colIdx, cell, entity);
                        colIdx++;
                    }
                    processFillData(entity);
                }
                log("Finish reading file");
                file.close();
                processClose();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        /**
         * Excel version 2003
         *
         * @param pathToFile
         */
        public void readExcelTypeXLS(String pathToFile) {
            try {
                FileInputStream file = new FileInputStream(new File(pathToFile));
                HSSFWorkbook workbook = new HSSFWorkbook(file);
                HSSFSheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();
                Row row;
                Cell cell;
                while (rowIterator.hasNext()) {
                    row = rowIterator.next();
                    Iterator<Cell> cellIterator = row.cellIterator();
                    int colIdx = 0;

                    AsiaDJ entity = new AsiaDJ();
                    while (cellIterator.hasNext()) {
                        cell = cellIterator.next();
                        processAddCellValue(colIdx, cell, entity);
                        colIdx++;
                    }
                    processFillData(entity);
                }
                log("Finish reading file");
                file.close();
                processClose();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

        private void processAddCellValue(int colIdx, Cell cell, AsiaDJ entity) {
            switch (colIdx) {
                case 0:
                    entity.setEmail(cell.getStringCellValue());
                    break;
                case 1:
                    entity.setFirstName(cell.getStringCellValue());
                    break;
                case 2:
                    entity.setLastName(cell.getStringCellValue());
                    break;
                case 3:
                    entity.setCountry(cell.getStringCellValue());
                    break;
                case 4:
                    entity.setDjName01(cell.getStringCellValue());
                    break;
                case 5:
                    entity.setDjName02(cell.getStringCellValue());
                    break;
                case 6:
                    entity.setDjName03(cell.getStringCellValue());
                    break;
                case 7:
                    entity.setDjName04(cell.getStringCellValue());
                    break;
                case 8:
                    entity.setDjName05(cell.getStringCellValue());
                    break;
            }
        }
    }

    private void log(String message) {
        log(message, null);
    }

    private void log(String message, Exception e) {
        String messageText = message + (e == null ? "" : e.getMessage());
        txaLog.setText(txaLog.getText() + messageText + "\n");
        if (e != null) {
            log.error(messageText, e);
        } else {
            log.info(messageText);
        }
    }

    /**
     * PROCESS WITH SELENIUM
     */
    private void processFillData(AsiaDJ entity) {
        log("Process fill data onto form");
        try {
            waitForSecond(500);
            processOpenChrome();
            driver.navigate().to(URL_TO_FORM);
            WebElement body = driver.findElement(By.cssSelector("body"));
            if (body.isDisplayed()) {

                List<WebElement> inputList = driver.findElements(By.cssSelector("form input"));
                for (int i = 0; i < inputList.size(); i++) {
                    switch (i) {
                        case 0:
                            inputList.get(i).sendKeys(entity.getEmail());
                            break;
                        case 1:
                            inputList.get(i).sendKeys(entity.getFirstName());
                            break;
                        case 2:
                            inputList.get(i).sendKeys(entity.getLastName());
                            break;
                        case 3:
                            inputList.get(i).sendKeys(entity.getCountry());
                            break;
                        case 4:
                            inputList.get(i).sendKeys(entity.getDjName01());
                            break;
                        case 5:
                            inputList.get(i).sendKeys(entity.getDjName02());
                            break;
                        case 6:
                            inputList.get(i).sendKeys(entity.getDjName03());
                            break;
                        case 7:
                            inputList.get(i).sendKeys(entity.getDjName05());
                            break;
                        case 8:
                            inputList.get(i).sendKeys(entity.getDjName05());
                            break;
                    }
                    waitForSecond(500);
                }
                // Submit
                WebElement submitTag = driver.findElement(By.xpath("//*[@id=\"mG61Hd\"]/div/div[2]/div[3]/div[3]/div/div"));
                submitTag.click();
                waitForSecond(25000);
            }
        } catch (Exception e) {
            log("Unable to find Element", e);
        }
    }

    private void processOpenChrome() {
        log("Opening Chrome...");
        try {
            driver.get("http://www.google.com/xhtml");
        } catch (Exception e) {
            log("Invalid Chrome Driver", e);
            return;
        }
    }

    private void processClose() {
        driver.close();
        driver.quit();
    }

    private void waitForSecond(long ms) {
        try {
            Thread.sleep(ms);
        } catch (Exception e) {
            log("Waiting face with issue", e);
        }
    }
}
