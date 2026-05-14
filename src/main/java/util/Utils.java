package util;

public class Utils {

	public String addTipoAtributo(String variavel, int valor) {
		if(variavel.contains("-")||variavel.contains("/")) {
			return "DATE,";
		}else if (variavel.contains(",") || variavel.contains(".")) {
			return "FLOAT,";
		} else if (variavel!=null && valor >0) {
			return "VARCHAR("+String.valueOf(valor)+")";
		}
		return null;
	}

	public String limpaQuery(String query) {
		return query.replace("[", " ").replace("]", "").replace(", ", " ");
	}
}
