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

	private static SXSSFWorkbook workbook;
	private static Sheet sheet;
	private static Integer row = 0;
	private static Integer iteracao = 1;
	private static String arquivoAtual;
	private static final String EXTENSAO_ARQUIVO = ".xlsx";
	private static final String REGEX_EXTRAI_NUMERO = "\\D";
	private static final String PREFIXO_ARQUIVO = "experimento";
	private static final String PREFIXO_ABA = "iteracao";
	private static final String DIRETORIO = "resources/xlsx/";

	public static void novoArquivoXls() {
		workbook = new SXSSFWorkbook(100);
		arquivoAtual = DIRETORIO + próximoArquivo();
	}

	// pega o nome do arquivo com a maior sequencia dos experimentos já
	// executados, e determina o nome do proximo arquivo com a proxima sequencia
	private static String próximoArquivo() {
		File diretorio = new File(DIRETORIO);

		List<File> files = Arrays.asList(diretorio
				.listFiles((dir, name) -> name.endsWith(EXTENSAO_ARQUIVO)));
		int experimento = 1;
		if (files != null) {
			files.sort((s1, s2) -> Long.compare(
					Long.parseLong(s1.getName().replaceAll(REGEX_EXTRAI_NUMERO,
							"")),
					Long.parseLong(s2.getName().replaceAll(REGEX_EXTRAI_NUMERO,
							""))));
			experimento = (Integer.parseInt(files.get(files.size() - 1)
					.getName().replaceAll(REGEX_EXTRAI_NUMERO, "")) + 1);

		}

		return PREFIXO_ARQUIVO + experimento + EXTENSAO_ARQUIVO;
	}

	public static void inserirAba() {
		sheet = workbook.createSheet(PREFIXO_ABA + (iteracao++));
		row = 0;
		Row newRow = sheet.createRow(row++);
		newRow.createCell(0).setCellValue("IMUNES");
		newRow.createCell(1).setCellValue("PSEUDO-IMUNE");
		newRow.createCell(2).setCellValue("INFECTANTES GERADOS");
		newRow.createCell(3).setCellValue("DOENTES");
		newRow.createCell(4).setCellValue("ACIDENTADOS");
		newRow.createCell(5).setCellValue("SADIOS");
		newRow.createCell(6).setCellValue("NASCIMENTOS");
	}

	public static void tratarResultado(Resultados resultado) {
		Row newRow = sheet.createRow(row++);
		newRow.createCell(0).setCellValue(resultado.getNumeroImunes());
		newRow.createCell(1).setCellValue(resultado.getNumeroPseudoImunes());
		newRow.createCell(2).setCellValue(
				resultado.getNumeroInfectantesGerados());
		newRow.createCell(3).setCellValue(resultado.getNumeroDoentes());
		newRow.createCell(4).setCellValue(resultado.getNumeroAcidentados());
		newRow.createCell(5).setCellValue(resultado.getNumeroSadios());
		newRow.createCell(6).setCellValue(resultado.getNascimentos());
	}

	public static void inserirDados() {
		try (FileOutputStream out = new FileOutputStream(arquivoAtual)) {
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
