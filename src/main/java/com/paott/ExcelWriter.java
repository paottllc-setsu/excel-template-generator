package com.paott;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

public class ExcelWriter {

	public static void main(String[] args) {
		// args[]引数にデータが渡されない時の処理
		if (args.length < 2) {
			System.err.println("Usage: java ExcelWriter templateFilePath outputFilePath");
			System.exit(1); // エラーの場合プログラムを終了する
		}
		String templateFilePath = args[0]; // テンプレートファイル
		String outputFilePath = args[1];  // アウトプットファイル
		// 受け取ったデータをJSON設定ファイルを参照して受け取る
		String cellConfigFilePath = "cellConfig.json";
		FileInputStream fis = null;
		Workbook workbook = null;
		
		try {
			fis = new FileInputStream(templateFilePath);
			workbook = new XSSFWorkbook(fis);
			Sheet sheet = workbook.getSheetAt(0); // 最初のシートを取得
			// 設定ファイルを読み込む
			JSONObject cellConfig = loadJsonFromFile(cellConfigFilePath);
			
			// 受け取ったデータ標準入力からJSONデータを読み込む
			BufferedReader reader = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8));
			StringBuilder sb = new StringBuilder();
			String line;
			while ((line = reader.readLine()) != null) {
				sb.append(line);
			}
			String jsonData = sb.toString();
			JSONObject jsonObject = new JSONObject(jsonData);
			
			// VB.NETからexecModeの値を受け取る mitumori Or seikyu
			String mode = jsonObject.getString("mode");
			
			// modeの値がnullまたは空文字列の場合、エラー処理を行う
			if (mode == null || mode.isEmpty()) {
				System.err.println("modeの値が設定されていません。");
				System.exit(5);
				return;
			}
			
			// 受け取ったmodeの値に基づいて、読み込むJSON設定ファイルを切り替える
			// modeが増えれば、else if ("mode2".equals(mode)) で条件式を増やし、jsonファイルを読み込ませる
			JSONArray modeConfig;
			if ("mode0".equals(mode)) {
				modeConfig = loadJsonArrayFromFile("mode0Config.json");
			} else if ("mode1".equals(mode)) {
				modeConfig = loadJsonArrayFromFile("mode1Config.json");
			} else {
				// エラー処理
				System.err.println("予期しない mode の値です。");
				System.exit(5);
				return;
			}

			// configJsonがnullの場合、エラー処理を行う
			if (modeConfig == null) {
				System.err.println(modeConfig + " の読み込みに失敗しました。");
				System.exit(5);
				return;
			}
			
			// データの書き込み処理を記述
			if (sheet != null) {
				writeData(workbook, sheet, jsonObject, cellConfig, modeConfig); // modeConfigをwriteDataに渡す
			}
			FileOutputStream fops = new FileOutputStream(outputFilePath);
			workbook.write(fops);
			fops.close();
			
			System.out.println("Excelファイルの書き込みに成功しました。");
			System.out.flush(); // 明示的にフラッシュ
			System.exit(0);
		} catch(IOException e) {
			System.err.println("エラーが発生しました。：" + e.getMessage());
			e.printStackTrace();
			System.out.flush(); // 明示的にフラッシュ
			System.exit(2);
		} catch(JSONException e) {
			System.err.println("JSONデータの解析に失敗しました。：" + e.getMessage());
			e.printStackTrace();
			System.exit(3);
		} finally {
			try {
				if (workbook != null) {
					workbook.close();
				}
				if (fis != null) {
					fis.close();
				}
			} catch(IOException ex) {
				ex.printStackTrace();
				System.exit(4);
			}
		}
	}

	private static JSONObject loadJsonFromFile(String filePath) throws IOException, JSONException {
		ClassLoader classLoader = ExcelWriter.class.getClassLoader();
		InputStream inputStream = classLoader.getResourceAsStream(filePath);
		if (inputStream == null) {
			throw new IOException("設定ファイルが見つかりません: " + filePath);
		}
		String content = new String(inputStream.readAllBytes(), StandardCharsets.UTF_8);
		return new JSONObject(content);
	}
	
	public static JSONArray loadJsonArrayFromFile(String filePath) throws IOException {
		ClassLoader classLoader = ExcelWriter.class.getClassLoader();
		InputStream inputStream = classLoader.getResourceAsStream(filePath);
		if (inputStream == null) {
			throw new IOException("設定ファイルが見つかりません: " + filePath);
		}
		String content = new String(inputStream.readAllBytes(), StandardCharsets.UTF_8);
		return new JSONArray(content);
	}
	
	public static void writeData(Workbook workbook, Sheet sheet, JSONObject jsonObject, JSONObject cellConfig, JSONArray modeConfig) {
		// mode.jsonのデータを処理
		for (int i = 0; i < modeConfig.length(); i++) {
			JSONObject config = modeConfig.getJSONObject(i);
			String cellAddress = config.getString("cellAddress");
			String value = config.getString("value"); // valueキーは必須
			if (cellAddress != null && !cellAddress.isEmpty()) {
				try {
					writeCell(sheet, cellAddress, value);
				} catch (IllegalArgumentException e) {
					System.err.println("セルアドレスが無効です。key:" + config.getString("key") + " address:" + cellAddress);
				}
			}
		}
		// 渡されたJSONデータの書き込み処理
		for (String key : JSONObject.getNames(cellConfig)) {
			if (key.equals("items")) { // itemsの処理
				if (jsonObject.has("items")) {
					JSONArray items = jsonObject.getJSONArray("items");
					JSONObject itemsConfig = cellConfig.getJSONObject("items");
					writeItems(workbook, sheet, items, itemsConfig);
				} else {
					System.err.println("JSONデータにitemsキーが存在しません。");
				}
			} else { //その他のキーの処理
				String cellAddress = cellConfig.optString(key);
				String value = jsonObject.optString(key);
				if (cellAddress != null && !cellAddress.isEmpty()) {
					try {
						writeCell(sheet, cellAddress, value);
					} catch (IllegalArgumentException e) {
						System.err.println("セルアドレスが無効です。key:" + key + "address:" + cellAddress);
					}
				}
			}
		}
	}

	// items配列の各要素をExcelシート (sheet) に書き込むためのメソッド
	private static void writeItems(Workbook workbook, Sheet sheet, JSONArray itemArray, JSONObject items) {
		int page1StartRow = items.getInt("page1startRow"); // cellConfigの開始行 (17行目である16) を取得 startRow は書き始めるExcelシートの行番号
		int page2StartRow = items.getInt("page2StartRow"); // 2ページ目の開始行 設定ファイルから取得（40行目である39）
		int page3StartRow = items.getInt("page3StartRow"); // 3ページ目の開始行 設定ファイルから取得（76行目である75）
		int page1Items = items.getInt("page1Items"); // 1ページ目は21項目を表示
		int page2Items = items.getInt("page2Items"); // 2ページ目は34項目を表示
		int page3Items = items.getInt("page3Items"); // 3ページ目は33項目を表示
		int totalAmount = 0; // H列の合計値を格納する変数
		int itemsLength = itemArray.length(); // JSONArrayの長さ
		int writtenItems = 0;
		int printEndRow = 0;
		
		// 1ページ目の書き込み
		//for (int i = 0; i < items.length(); i++) { // itemsというJASONArrayの長さだけループする
		for (int i = 0; i < Math.min(itemsLength, page1Items); i++) {
			int rowNum = page1StartRow + i; // 行番号
			JSONObject item = itemArray.getJSONObject(i); // items配列から i 番目の要素を取得
			writeItemRow(sheet, item, items, rowNum);
			// H列の値を合計
			try {
				totalAmount += Integer.parseInt(item.getString("column5"));
			} catch (NumberFormatException e) {
				// 数値に変換できない場合は無視
			}
			writtenItems++;
		}
		printEndRow = page1StartRow + page1Items;
		// 2ページ目の書き込み
		if (itemsLength > page1Items) {
			for (int i = 0; i < Math.min(itemsLength - page1Items, page2Items); i++) {
				int rowNum = page2StartRow + i;
				JSONObject item = itemArray.getJSONObject(page1Items + i);
				writeItemRow(sheet, item, items, rowNum);
				try {
					totalAmount += Integer.parseInt(item.getString("column5"));
				} catch (NumberFormatException e) {
					// 数値に変換できない場合は無視
				}
				writtenItems++;
			}
			printEndRow = page2StartRow + page2Items;
		}
		// 3ページ目の書き込み
		if (itemsLength > page1Items + page2Items) {			
			for (int i = 0; i < Math.min(itemsLength - page1Items - page2Items, page3Items); i++) {
				int rowNum = page3StartRow + i;
				JSONObject item = itemArray.getJSONObject(page1Items + page2Items + i);
				writeItemRow(sheet, item, items, rowNum);
				try {
					totalAmount += Integer.parseInt(item.getString("column5"));
				} catch (NumberFormatException e){
					// 数値に変換できない場合は無視
				}
				writtenItems++;
			}
			printEndRow = page3StartRow + page3Items;
		}
		// 合計行の書き込み
		if (writtenItems > 0) {
			int totalRow;
			if (writtenItems < page1Items) { // totalRow含めて1ページ内(37行目)で収まる時
				totalRow = page1StartRow + page1Items; // 37行目
			} else if (writtenItems == page1Items) { // 1ページ目の最後(37行目)にデータが入っている時
				totalRow = page2StartRow + page2Items; // 73行目
			} else if (writtenItems < page1Items + page2Items) { // totalRow含めて2ページ内(73行目)で収まる時
				totalRow = page2StartRow + page2Items; // 73行目
			} else { 
				totalRow = page3StartRow + page3Items; // その他は全て109行目
			}
			//writeCell(sheet, "B" + totalRow, "合計"); 
			//writeCell(sheet, "H" + totalRow, String.valueOf(totalAmount));
			writeCell(sheet, items.getString("column1") + totalRow, "合計");
			writeCell(sheet, items.getString("column5") + totalRow, String.valueOf(totalAmount));
		}
		workbook.setPrintArea(0, "$A$1:$M$" + printEndRow);
	}
	
	private static void writeItemRow(Sheet sheet, JSONObject item, JSONObject items, int rowNum) {
		Row row = sheet.getRow(rowNum); // Excelシートの startRow + i 行目のオブジェクトを取得し row に代入
		if (row == null) {
			row = sheet.createRow(rowNum); // 取得しようとした行が null の場合、新しい行を作成し row に代入
		}
		// writeCell はExcelシートの指定されたセルに値を書き込む関数 
		// itemsConfig.getString("column") 設定ファイルから"column"に対応する列名(B列)を取得
		// item.getString("column") 現在の"column"に対応する値を取得
		writeCell(sheet, items.getString("column1") + (rowNum + 1), item.getString("column1"));
		writeCell(sheet, items.getString("column2") + (rowNum + 1), item.getString("column2"));
		writeCell(sheet, items.getString("column3") + (rowNum + 1), item.getString("column3"));
		writeCell(sheet, items.getString("column4") + (rowNum + 1), item.getString("column4"));
		writeCell(sheet, items.getString("column5") + (rowNum + 1), item.getString("column5"));
		writeCell(sheet, items.getString("column6") + (rowNum + 1), item.getString("column6"));
	}

	public static void writeCell(Sheet sheet, String cellAddress, String cellvalue) { // cellAddress 書き込み先のセルアドレス ("A1", "B5")を表す文字列
		if (cellAddress == null || cellAddress.isEmpty()) {
			System.err.println("セルアドレスが無効です。");
		}
		try {
			CellReference cellReference = new CellReference(cellAddress); // CellReferenceクラス セルアドレスを解析し、行番号と列番号を取得するためのクラス
			int rowIndex = cellReference.getRow(); // セルアドレスから行番号を取得し、rowIndexに格納
			int columnIndex = cellReference.getCol(); // セルアドレスから列番号を取得し、columnIndexに格納
			Row row = sheet.getRow(rowIndex); // 指定された行番号の行オブジェクトを取得し、rowに格納
			if (row == null) { // row が null の時
				row = sheet.createRow(rowIndex); // 新しい行を作成し、rowに格納
			}

			Cell cell = row.getCell(columnIndex); // 指定された列番号の列オブジェクトを取得し、cellに格納
			if (cell == null) { // cell が null の時
				cell = row.createCell(columnIndex); // 新しいセルを作成し、cellに格納
			}
			// 整数として解釈できるか判定
			try {
				int numericValue = Integer.parseInt(cellvalue);
				cell.setCellValue(numericValue);
			} catch (NumberFormatException e) {
				// 整数として解釈できない場合は文字列として書き込み
				cell.setCellValue(cellvalue);
			}
		} catch (IllegalArgumentException e) {
			System.err.println("セルアドレスが無効です。 address:" + cellAddress);
			throw e;
		}
	}
}