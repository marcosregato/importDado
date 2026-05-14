package rascunho;

import com.codoid.products.fillo.Connection;
import com.codoid.products.fillo.Fillo;
import com.codoid.products.fillo.Recordset;

import java.util.ArrayList;
import java.util.List;

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
                if (connection != null) connection.close();
            } catch (Exception e) {
                    // Ignore close errors
            }
        }
    }

    public static void main(String[] args) {
        SqlExcel sqlExcel = new SqlExcel();
        System.out.println(sqlExcel.buscarAba_1("banana"));
        sqlExcel.updateAba_1("laranja");
        sqlExcel.deleteAba_1("morango", "0");
        List<String> resultados = sqlExcel.listSelectAba_1();
        for (int x = 0; x < resultados.size(); x++) {
            System.out.println("Linha => " + resultados.get(x));
        }
    }
}
