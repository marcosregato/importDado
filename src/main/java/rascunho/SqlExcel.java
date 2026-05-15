package rascunho;

import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;



import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import java.io.File;

public class SqlExcel {

    private static final String ABA_1 = "ABA_1";
    private final String pathExcel = "src/teste.xls";

    public List<String> listSelectAba_1() {
        List<String> lista = new ArrayList<>();
        Connection connection = null;
        Recordset recordset = null;
        try {
            Fillo fillo = new Fillo();
            connection = fillo.getConnection(pathExcel);
            String query = "select * from " + ABA_1;
            recordset = connection.executeQuery(query);

            while (recordset.next()) {
                String ativo = recordset.getField("ativo");
                if (ativo != null && !ativo.isEmpty() && !ativo.equals("0")) {
                    lista.add(recordset.getField("fruta"));
                    lista.add(recordset.getField("nome"));
                    lista.add(recordset.getField("cidade"));
                }
            }
        } catch (Exception e) {
            System.err.println("Erro na listSelectAba_1: " + e.getMessage());
        } finally {

            if (recordset != null) recordset.close();
            if (connection != null) connection.close();
        }
        return lista;
    }

    public void updateAba_1(String nome) {
        Connection connection = null;
        try {
            Fillo fillo = new Fillo();
            connection = fillo.getConnection(pathExcel);
            String query = "update " + ABA_1 + " set fruta = '" + nome + "'" + " where cidade = 'sp'";
            connection.executeUpdate(query);
        } catch (Exception e) {
            System.err.println("Erro na updateAba_1: " + e.getMessage());
        } finally {
            try {
                if (connection != null) connection.close();
            } catch (Exception e) {
                // Ignore close errors
            }
        }
    }

    public List<String> buscarAba_1(String valor) {
        List<String> lista = new ArrayList<>();
        Connection connection = null;
        Recordset recordset = null;
        try {
            Fillo fillo = new Fillo();
            connection = fillo.getConnection(pathExcel);
            String query = "select * from " + ABA_1 + " where fruta = '" + valor + "'";
            recordset = connection.executeQuery(query);
            while (recordset.next()) {
                lista.add(recordset.getField("fruta"));
                lista.add(recordset.getField("nome"));
                lista.add(recordset.getField("cidade"));
            }
        } catch (Exception e) {
            System.err.println("Erro na buscarAba_1: " + e.getMessage());
        } finally {
            try {
                if (recordset != null) recordset.close();
                if (connection != null) connection.close();
            } catch (Exception e) {
                // Ignore close errors
            }
        }
        return lista;
    }

    public void deleteAba_1(String valor, String status) {
        Connection connection = null;
        try {
            Fillo fillo = new Fillo();
            connection = fillo.getConnection(pathExcel);
            String query = "update " + ABA_1 + " set ATIVO = '"+status+"' where fruta = '" + valor + "'";
            connection.executeUpdate(query);
        } catch (Exception e) {
            System.err.println("Erro na deleteAba_1: " + e.getMessage());
        } finally {
            try {
                if (connection != null) {
                    connection.close();
                }
            } catch (Exception e) {
                    // Ignore close errors
                    e.printStackTrace();
            }
        }
    }

    

    //TODO: A ultima coluna da planilha do excel da aba 1
    public void getUltimaColunaPlanilha(){
        try {
            FileInputStream arq = new FileInputStream(new File(pathExcel));
            Workbook workbook = new HSSFWorkbook(arq);
            Sheet sheet = workbook.getSheetAt(0);

            Row row = sheet.getRow(0);

            if(row != null){
                int valorDaUltimaColuna = row.getLastCellNum() - 1;
                Cell nomeDaUltimaColuna = row.getCell(valorDaUltimaColuna);
                System.out.println("Nome da coluna " + nomeDaUltimaColuna.getStringCellValue());
                System.out.println("Indece da coluna: " + valorDaUltimaColuna);
            }

            workbook.close();
            arq.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /*
     * Adicionar duas colunas no final da ABA_1 com nome "ID" e "ATIVO"
     * Lembrando que na coluna ATIVO deve receber o valor 0(zero) ou 1(um)
     * a coluna ID deve ser incrementada automaticamente
     */
    public void adicionarColunaIdEAtivo() {
        FileInputStream arq = null;
        Workbook workbook = null;
        FileOutputStream escreverArquivo = null;
        try {
            arq = new FileInputStream(new File(pathExcel));
            workbook = new XSSFWorkbook(arq);


            Sheet sheet = workbook.getSheetAt(0);
            //int linha = sheet.getLastRowNum();

            Row row = sheet.getRow(0);
            if (row != null){
                row= sheet.createRow(0);
            }

            int contColuna = row.getLastCellNum();
            Cell cellID = row.createCell(contColuna);
            cellID.setCellValue("ID");

            Cell cellAtivo = row.createCell(contColuna + 1);
            cellAtivo.setCellValue("ATIVO");

            escreverArquivo = new FileOutputStream(pathExcel);
            workbook.write(escreverArquivo);
            System.out.println("Colunas adicionadas com sucesso!");

        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (escreverArquivo != null) escreverArquivo.close();
                if (workbook != null) workbook.close();
                if (arq != null) arq.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }

    /**
     * Na coluna ATIVO deve ter o valor 1, para todos os registros
     */
    public void setValorAtivo() {

    }


    /**  
     * criar um metodo criar um coluna no excel para ser ID da linha e incrementar 
     * automaticamente
    */
    public int crarIdDaLinha() {

        return 0;
    }

    /**
     * Esse coluna é a Primary Key da tabela SQL 
     * @return
     */
    public int getValorIdDaLinha(String nomeDaTabela) {
        return 0;
    }


    /**
     * Essa coluna é Foriegn Key da tabela SQL
     */
    public int getValorForeignKeyDaTabela() {
        return 0;
    }

    /**
     * Talvez precise implementar Árvore B+ (B+ Tree) 
     * é uma estrutura de dados baseada em árvores de busca
     * inserção com divisão de nó (split)
     * Relacionamento 1:1
     */
    public void relacionamentoUmParaUm() {

    }

    /**
     * Relacionamento 1:N
     */
    public void relacionamentoUmParaMuitos() {

    }

    /*
    Posso usar o metodo salvarAba_1 para salvar um objeto Aba1(nome da tabela SQL) no excel
    public void salvarAba_1(Aba1 aba1) {
        
        try{
            Fillo fillo = new Fillo();
            Connection connection = fillo.getConnection(pathExcel);
            String insert = "INSERT INTO " + ABA_1 + " (fruta, nome, cidade, ativo) VALUES ('" + aba1.getFruta() + "', '" + aba1.getNome() + "', '" + aba1.getCidade() + "', '" + aba1.getAtivo() + "')";
            connection.executeUpdate(insert);
        }catch(Exception e){
            System.err.println("Erro na salvarAba_1: " + e.getMessage());
        }
    }
    */

    public void salvarAba_1(String fruta, String nome, String cidade, String ativo) {
        try{
            Fillo fillo = new Fillo();
            Connection connection = fillo.getConnection(pathExcel);
            String insert = "INSERT INTO " + ABA_1 + " (fruta, nome, cidade, ativo) VALUES ('"+fruta+"', '"+nome+"', '"+cidade+"', '"+ativo+"')";
            connection.executeUpdate(insert);
        }catch(Exception e){
            System.err.println("Erro na salvarAba_1: " + e.getMessage());
        }
    }

    public static void main(String[] args) {
        SqlExcel sqlExcel = new SqlExcel();
        
        
        /*System.out.println(sqlExcel.buscarAba_1("banana"));
        sqlExcel.updateAba_1("laranja");
        sqlExcel.deleteAba_1("morango", "0");
        List<String> resultados = sqlExcel.listSelectAba_1();
        for (int x = 0; x < resultados.size(); x++) {
            System.out.println("Linha => " + resultados.get(x));
        }*/
        sqlExcel.getUltimaColunaPlanilha();
    }
}
