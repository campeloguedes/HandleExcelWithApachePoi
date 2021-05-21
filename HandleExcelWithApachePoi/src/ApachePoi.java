import java.io.File;
import java.io.IOException;

public class ApachePoi {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		File file = new File("C:\\Users\\Leonardo\\git\\repository5\\HandleExcelWithApachePoi\\src\\arquivo_excel.xls");
		
		if (!file.exists()) {
			file.createNewFile();
		}
		
		
		
	}

}
