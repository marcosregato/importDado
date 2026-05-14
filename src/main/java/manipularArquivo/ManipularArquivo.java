package manipularArquivo;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import util.ConfigManager;

public class ManipularArquivo {

	public List<String> getListNomeAba(int quantidadeAba) {
		List<String> listNomeAba = new ArrayList<>();
		String excelPath = ConfigManager.getFullExcelPath();
		FileInputStream fs = null;
		Workbook wb = null;
		
		try {
			fs = new FileInputStream(excelPath);
			wb = Workbook.getWorkbook(fs);
			
			for(int aba = 0; aba <= quantidadeAba; aba++) {
				Sheet sh = wb.getSheet(aba);
				if(sh != null && sh.getName() != null && !sh.getName().trim().isEmpty()) {
					listNomeAba.add(sh.getName());
				}
			}
			
		} catch (IOException | BiffException e) {
			System.err.println("Erro ao ler abas do Excel: " + e.getMessage());
			e.printStackTrace();
		} finally {
			try {
				if (wb != null) {
					wb.close();
				}
				if (fs != null) {
					fs.close();
				}
			} catch (IOException e) {
				System.err.println("Erro ao fechar recursos: " + e.getMessage());
			}
		}
		
		return listNomeAba;
	}

	public List<String> getListColuna(){
		List<String> listColuna = new ArrayList<>();
		String excelPath = ConfigManager.getFullExcelPath();
		String sheetName = ConfigManager.getDefaultSheetName();
		FileInputStream fs = null;
		Workbook wb = null;
		
		try {
			fs = new FileInputStream(excelPath);
			wb = Workbook.getWorkbook(fs);
			Sheet sh = wb.getSheet(sheetName);
			
			if (sh == null) {
				System.err.println("Aba '" + sheetName + "' não encontrada no arquivo Excel.");
				return Collections.emptyList();
			}
			
			// Lê cabeçalhos da primeira linha (até máximo de colunas)
			int maxColumns = Math.min(ConfigManager.getMaxColumnsToRead(), sh.getColumns());
			for(int x = 0; x < maxColumns; x++) {
				String celula = sh.getCell(x, 0).getContents();
				if(celula != null && !celula.trim().isEmpty()) {
					listColuna.add(celula.trim());
				}
			}
			
		} catch (Exception e) {
			System.err.println("Erro ao ler colunas do Excel: " + e.getMessage());
			e.printStackTrace();
		} finally {
			try {
				if (wb != null) {
					wb.close();
				}
				if (fs != null) {
					fs.close();
				}
			} catch (IOException e) {
				System.err.println("Erro ao fechar recursos: " + e.getMessage());
			}
		}
		
		return listColuna;
	}

}
