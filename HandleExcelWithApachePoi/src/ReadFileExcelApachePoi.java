import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class ReadFileExcelApachePoi {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		FileInputStream entrada = new FileInputStream(new File("C:\\Users\\Leonardo\\git\\repository5\\HandleExcelWithApachePoi\\src\\arquivo_excel.xls"));
		
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(entrada); /*prepara a entrada do arquivo excel para leitura*/
		HSSFSheet planilha = hssfWorkbook.getSheetAt(0); /*Pega a primeira planilha do arquivo excel*/
		
		Iterator<Row> linhaIterator = planilha.iterator();
		
		List<Pessoa> pessoas = new ArrayList<Pessoa>();
		
		while (linhaIterator.hasNext()) { /*enquato existir linha ele vai ler a linha*/
			Row linha = linhaIterator.next(); /*dados da linha*/
			
			Iterator<Cell> celulas = linha.cellIterator();
			
			Pessoa pessoa = new Pessoa();
			while (celulas.hasNext()) {
				Cell cell = celulas.next();
				
				switch (cell.getColumnIndex()) {
				case 0:
					pessoa.setNome(cell.getStringCellValue());
					break;
				case 1:
					pessoa.setEmail(cell.getStringCellValue());
					break;
				case 2:
					pessoa.setIdade(Double.valueOf(cell.getNumericCellValue()).intValue());
					break;

				}
					
			}
			
			pessoas.add(pessoa);
			
		}
		
		entrada.close();
		
		for (Pessoa p : pessoas) {
			System.out.println(p);
		}
		
		
		
		

	}

}
