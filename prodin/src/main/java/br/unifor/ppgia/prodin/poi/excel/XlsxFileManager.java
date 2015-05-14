package br.unifor.ppgia.prodin.poi.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import br.unifor.ppgia.prodin.atualizacoes.Resultados;

public class XlsxFileManager {

	private static SXSSFWorkbook workbook = new SXSSFWorkbook(100);
	private static Sheet sheet;
	private static Integer row = 0;
	private static Integer iteracao = 1;
	private static String arquivoAtual = nomeProximoArquivo();

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

	public static void createSheet() {
		sheet = workbook.createSheet("iteracao" + (iteracao++));
	}

	public static void inserirResultado(Resultados resultado) {
		Row newRow = sheet.createRow(row++);
		newRow.createCell(0).setCellValue(resultado.getNumeroImunes());
		newRow.createCell(1).setCellValue(resultado.getNumeroPseudoImunes());
		newRow.createCell(2).setCellValue(
				resultado.getNumeroInfectantesGerados());
		newRow.createCell(3).setCellValue(resultado.getNumeroDoentes());
		newRow.createCell(4).setCellValue(resultado.getNumeroAcidentados());
		newRow.createCell(5).setCellValue(resultado.getNumeroSadios());
		newRow.createCell(6).setCellValue(resultado.getNascimentos());

		try (FileOutputStream out = new FileOutputStream(arquivoAtual)) {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void main(String[] args) {
		System.out.println(XlsxFileManager.nomeProximoArquivo());
	}

}
