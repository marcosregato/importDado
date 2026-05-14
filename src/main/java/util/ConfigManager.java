package util;

import java.io.IOException;
import java.io.InputStream;
import java.util.Properties;

public class ConfigManager {
    private static Properties properties;
    private static final String CONFIG_FILE = "config.properties";
    
    static {
        loadProperties();
    }
    
    private static void loadProperties() {
        properties = new Properties();
        try (InputStream input = ConfigManager.class.getClassLoader().getResourceAsStream(CONFIG_FILE)) {
            if (input == null) {
                throw new RuntimeException("Arquivo de configuração não encontrado: " + CONFIG_FILE);
            }
            properties.load(input);
        } catch (IOException e) {
            throw new RuntimeException("Erro ao carregar configuração", e);
        }
    }
    
    public static String getBasePath() {
        return properties.getProperty("base.path", "src/");
    }
    
    public static String getExcelFile() {
        return properties.getProperty("excel.file", "planilha_teste.xls");
    }
    
    public static String getOutputDirectory() {
        return properties.getProperty("output.directory", "arqSQL/");
    }
    
    public static String getDefaultSheetName() {
        return properties.getProperty("default.sheet.name", "TURMA");
    }
    
    public static String getFullExcelPath() {
        return getBasePath() + getExcelFile();
    }
    
    public static String getFullOutputPath() {
        return getBasePath() + getOutputDirectory();
    }
    
    public static int getMaxColumnsToRead() {
        return Integer.parseInt(properties.getProperty("max.columns.to.read", "50"));
    }
    
    public static int getMaxRowsToRead() {
        return Integer.parseInt(properties.getProperty("max.rows.to.read", "1000"));
    }
}
