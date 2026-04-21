package util;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;

/**
 * Classe unificada de leitura Excel que suporta tanto JXL quanto Apache POI
 */
public class ExcelReader {
    
    /**
     * Obtém nomes das abas usando biblioteca JXL
     */
    public List<String> getSheetNames(int maxSheets) {
        List<String> sheetNames = new ArrayList<>();
        String excelPath = ConfigManager.getFullExcelPath();
        FileInputStream fs = null;
        Workbook wb = null;
        
        try {
            fs = new FileInputStream(excelPath);
            wb = Workbook.getWorkbook(fs);
            
            int sheetCount = Math.min(maxSheets + 1, wb.getNumberOfSheets());
            for (int i = 0; i < sheetCount; i++) {
                Sheet sheet = wb.getSheet(i);
                if (sheet != null && sheet.getName() != null && !sheet.getName().trim().isEmpty()) {
                    sheetNames.add(sheet.getName().trim());
                }
            }
            
        } catch (IOException | BiffException e) {
            System.err.println("Erro ao ler nomes das abas: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                if (wb != null) wb.close();
                if (fs != null) fs.close();
            } catch (IOException e) {
                System.err.println("Erro ao fechar recursos: " + e.getMessage());
            }
        }
        
        return sheetNames;
    }
    
    /**
     * Obtém cabeçalhos das colunas da primeira linha usando biblioteca JXL
     */
    public List<String> getColumnHeaders(String sheetName) {
        List<String> headers = new ArrayList<>();
        String excelPath = ConfigManager.getFullExcelPath();
        FileInputStream fs = null;
        Workbook wb = null;
        
        try {
            fs = new FileInputStream(excelPath);
            wb = Workbook.getWorkbook(fs);
            Sheet sheet = wb.getSheet(sheetName);
            
            if (sheet == null) {
                System.err.println("Aba '" + sheetName + "' não encontrada.");
                return Collections.emptyList();
            }
            
            int maxColumns = Math.min(ConfigManager.getMaxColumnsToRead(), sheet.getColumns());
            for (int col = 0; col < maxColumns; col++) {
                String cellValue = sheet.getCell(col, 0).getContents();
                if (cellValue != null && !cellValue.trim().isEmpty()) {
                    headers.add(cellValue.trim());
                }
            }
            
        } catch (IOException | BiffException e) {
            System.err.println("Erro ao ler cabeçalhos: " + e.getMessage());
            e.printStackTrace();
        } finally {
            try {
                if (wb != null) wb.close();
                if (fs != null) fs.close();
            } catch (IOException e) {
                System.err.println("Erro ao fechar recursos: " + e.getMessage());
            }
        }
        
        return headers;
    }
    
    /**
     * Obtém todos os dados da aba usando Apache POI
     */
    public List<List<String>> getSheetData(String sheetName) {
        List<List<String>> data = new ArrayList<>();
        String excelPath = ConfigManager.getFullExcelPath();
        
        try (FileInputStream fis = new FileInputStream(excelPath);
             HSSFWorkbook workbook = new HSSFWorkbook(fis)) {
            
            HSSFSheet sheet = workbook.getSheet(sheetName);
            if (sheet == null) {
                System.err.println("Aba '" + sheetName + "' não encontrada.");
                return Collections.emptyList();
            }
            
            int maxRows = Math.min(ConfigManager.getMaxRowsToRead(), sheet.getPhysicalNumberOfRows());
            
            for (int rowIndex = 0; rowIndex < maxRows; rowIndex++) {
                HSSFRow row = sheet.getRow(rowIndex);
                if (row != null) {
                    List<String> rowData = new ArrayList<>();
                    int maxCols = Math.min(ConfigManager.getMaxColumnsToRead(), row.getPhysicalNumberOfCells());
                    
                    for (int colIndex = 0; colIndex < maxCols; colIndex++) {
                        HSSFCell cell = row.getCell(colIndex);
                        String cellValue = getCellValueAsString(cell);
                        rowData.add(cellValue);
                    }
                    data.add(rowData);
                }
            }
            
        } catch (IOException e) {
            System.err.println("Erro ao ler dados da planilha: " + e.getMessage());
            e.printStackTrace();
        }
        
        return data;
    }
    
    /**
     * Converte valor da célula para string
     */
    private String getCellValueAsString(HSSFCell cell) {
        if (cell == null) {
            return "";
        }
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }
    
    /**
     * Detecta tipo de dado para geração de coluna SQL
     */
    public String detectDataType(String value) {
        if (value == null || value.trim().isEmpty()) {
            return "VARCHAR(50)";
        }
        
        value = value.trim();
        
        // Verifica padrões de data
        if (value.contains("-") || value.contains("/")) {
            return "DATE";
        }
        
        // Verifica números decimais
        if (value.contains(",") || value.contains(".")) {
            try {
                Double.parseDouble(value.replace(",", "."));
                return "DECIMAL(10,2)";
            } catch (NumberFormatException e) {
                // Não é um número válido, trata como string
            }
        }
        
        // Verifica números inteiros
        try {
            Integer.parseInt(value);
            return "INT";
        } catch (NumberFormatException e) {
            // Não é um inteiro
        }
        
        // Verifica valores booleanos
        if (value.equalsIgnoreCase("true") || value.equalsIgnoreCase("false") ||
            value.equalsIgnoreCase("sim") || value.equalsIgnoreCase("não") ||
            value.equalsIgnoreCase("yes") || value.equalsIgnoreCase("no")) {
            return "BOOLEAN";
        }
        
        // Padrão para VARCHAR com comprimento apropriado
        int length = Math.min(255, Math.max(50, value.length() + 10));
        return "VARCHAR(" + length + ")";
    }
}
