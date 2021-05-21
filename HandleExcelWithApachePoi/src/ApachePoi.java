import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;



public class ApachePoi {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file = new File("C:\\Users\\Leonardo\\git\\repository5\\HandleExcelWithApachePoi\\src\\arquivo_excel.xls");
		
		if (!file.exists()) {
			file.createNewFile();
		}
		
		Pessoa p1 = new Pessoa();
		p1.setEmail("pessoa1@gmail.com");
		p1.setIdade(50);
		p1.setNome("Ricardo Edigio");
		
		Pessoa p2 = new Pessoa();
		p2.setEmail("pessoa2@gmail.com");
		p2.setIdade(40);
		p2.setNome("Marcos Tadeu");
		
		Pessoa p3 = new Pessoa();
		p3.setEmail("pessoa3@gmail.com");
		p3.setIdade(30);
		p3.setNome("Maria Julia");
		
		List<Pessoa> pessoas = new ArrayList<Pessoa>();
		pessoas.add(p1);
		pessoas.add(p2);
		pessoas.add(p3);
		
		
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(); /*escrever na planilha*/
		HSSFSheet linhasPessoa = hssfWorkbook.createSheet("Planilha de pessoas"); /*criar planilha*/
		
		int nlinha = 0;
		for (Pessoa p : pessoas) {
			Row linha = linhasPessoa.createRow(nlinha ++); /*criando a linha na planilha*/
			
			int celula = 0;
			
			Cell celNome = linha.createCell(celula ++); /*celula 1*/
			celNome.setCellValue(p.getNome());
			
			Cell celEmail = linha.createCell(celula ++); /*celula 2*/
			celEmail.setCellValue(p.getEmail());
			
			Cell celIdade = linha.createCell(celula ++); /*celula 3*/
			celIdade.setCellValue(p.getIdade());
			
		}
		
		FileOutputStream saida = new FileOutputStream(file);
		hssfWorkbook.write(saida);
		
		saida.flush();
		saida.close();
		
		System.out.println("Planilha Criada com Sucesso!");
		
	}

}
