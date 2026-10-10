package refatoraNumero;

public class DescobirNumero {

    public void descobrir(){
        int[] conjunto = new int[2];
        int[] conjuntoNumeo = {1,2,3,4,5,6,7,8,9};

        int soma =0;
        while (soma != 9){
            for (int i = 0; i < conjuntoNumeo.length; i++) {
                soma = conjuntoNumeo[i] + soma;
                System.out.println(soma);

            }
        }
    }
    public static void main(String[] args) {
        DescobirNumero d = new DescobirNumero();
        d.descobrir();
    }

}
