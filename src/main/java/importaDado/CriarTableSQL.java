package importaDado;

import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.List;

import util.ConfigManager;
import util.ExcelReader;

public class CriarTableSQL {
    
    private ExcelReader excelReader = new ExcelReader();

	public String criarArquivoSQL(String arquivo) {
		try {
			String outputPath = ConfigManager.getFullOutputPath();
			
			// Criar diretório de saída se não existir
			Files.createDirectories(Paths.get(outputPath));
			
			File arq = new File(outputPath + arquivo);
			if (arq.exists()) {
				System.out.println("ARQUIVO JA EXISTE: " + arq.getAbsolutePath());
				return arq.getAbsolutePath();
			} else {
				arq.createNewFile();
				System.out.println("ARQUIVO CRIADO: " + arq.getAbsolutePath());
				return arq.getAbsolutePath();
			}
		} catch (IOException e) {
			System.err.println("Erro ao criar arquivo SQL: " + e.getMessage());
			e.printStackTrace();
		}
		return null;
	}

	public void criarTabela(String pathArqSQL, List<String> listNomeTabela){
		try (FileWriter myWriter = new FileWriter(pathArqSQL)) {
			for(int x =0; x < listNomeTabela.size();x++) {
				String nomeTabela = listNomeTabela.get(x); 
				
				// Obtém cabeçalhos das colunas usando ExcelReader
				List<String> colunas = excelReader.getColumnHeaders(nomeTabela);
				if (colunas.isEmpty()) {
					System.err.println("Nenhuma coluna encontrada para a tabela: " + nomeTabela);
					continue;
				}
				
				// Obtém dados de amostra para detectar tipos de dados
				List<List<String>> sampleData = excelReader.getSheetData(nomeTabela);
				
				StringBuilder tabela = new StringBuilder();
				tabela.append("CREATE TABLE ").append(nomeTabela).append(" (\n");
				tabela.append("    id INT PRIMARY KEY NOT NULL AUTO_INCREMENT,\n");
				
				// Adiciona colunas com tipos de dados detectados
				for(int s =0; s < colunas.size();s++) {
					String nomeColuna = colunas.get(s);
					if (nomeColuna != null && !nomeColuna.trim().isEmpty()) {
						String dataType = detectDataTypeForColumn(nomeColuna, sampleData, s);
						tabela.append("    ").append(nomeColuna.trim()).append(" ").append(dataType);
						if (s < colunas.size() - 1) {
							tabela.append(",");
						}
						tabela.append("\n");
					}
				}
				
				tabela.append(");\n\n");
				
				myWriter.write(tabela.toString());
				System.out.println("Tabela criada: " + nomeTabela + " (" + colunas.size() + " colunas)");
			}
			
			System.out.println("Arquivo SQL gerado com sucesso: " + pathArqSQL);
		} catch (IOException e) {
			System.err.println("Erro ao criar tabelas SQL: " + e.getMessage());
			e.printStackTrace();
		}
	}
	
	/**
	 * Detecta tipo de dado para uma coluna baseado nos dados de amostra
	 */
	private String detectDataTypeForColumn(String columnName, List<List<String>> data, int columnIndex) {
		if (data.isEmpty() || columnIndex >= data.get(0).size()) {
			return "VARCHAR(50)";
		}
		
		// Amostra algumas linhas para detectar tipo de dado
		int sampleSize = Math.min(5, data.size());
		for (int i = 1; i < sampleSize; i++) { // Pula linha de cabeçalho
			if (columnIndex < data.get(i).size()) {
				String value = data.get(i).get(columnIndex);
				String detectedType = excelReader.detectDataType(value);
				if (!detectedType.equals("VARCHAR(50)")) {
					return detectedType;
				}
			}
		}
		
		return "VARCHAR(50)";
	}

	public static void main(String[] args) {
		try {
			System.out.println("Iniciando geração de SQL a partir do Excel...");
			System.out.println("Arquivo Excel: " + ConfigManager.getFullExcelPath());
			
			ExcelReader excelReader = new ExcelReader();
			CriarTableSQL sql = new CriarTableSQL();
			
			// Obtém nomes das abas do arquivo Excel
			List<String> nomesAbas = excelReader.getSheetNames(5); // Lê até 6 abas (0-5)
			if (nomesAbas.isEmpty()) {
				System.err.println("Nenhuma aba encontrada no arquivo Excel.");
				return;
			}
			
			System.out.println("Abas encontradas: " + nomesAbas);
			
			// Cria arquivo SQL
			String sqlFilePath = sql.criarArquivoSQL("tables.sql");
			if (sqlFilePath == null) {
				System.err.println("Falha ao criar arquivo SQL.");
				return;
			}
			
			// Gera tabelas
			sql.criarTabela(sqlFilePath, nomesAbas);
			
			System.out.println("Processo concluído com sucesso!");
			
		} catch (Exception e) {
			System.err.println("Erro durante a execução: " + e.getMessage());
			e.printStackTrace();
		}
	}

}
