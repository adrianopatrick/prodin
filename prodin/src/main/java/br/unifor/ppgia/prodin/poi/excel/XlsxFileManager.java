package br.unifor.ppgia.prodin.poi.excel;

import java.io.File;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class XlsxFileManager {

	private static SXSSFWorkbook workbook = new SXSSFWorkbook(100);

	// pega o nome do arquivo com a maior sequencia dos experimentos já
	// executados, e determina o nome do proximo arquivo com a proxima sequencia
	public static String nomeProximoArquivo() {
		File diretorio = new File("resources/xslx/");
		System.out.println(diretorio.getAbsolutePath());

		List<File> files = Arrays.asList(diretorio
				.listFiles((dir, name) -> name.endsWith(".xslx")));
		files.sort((s1, s2) -> Long.compare(
				Long.parseLong(s1.getName().replaceAll("\\D", "")),
				Long.parseLong(s2.getName().replaceAll("\\D", ""))));

		return "experimento"
				+ (Long.parseLong(files.get(files.size() - 1).getName()
						.replaceAll("\\D", "")) + 1) + ".xslx";
	}

	public static void main(String[] args) {
		System.out.println(XlsxFileManager.nomeProximoArquivo());
	}

}
