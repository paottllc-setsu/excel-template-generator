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

public class ExcelWriter2 {

	public static void main(String[] args) {
		// args[]引数にデータが渡されない時の処理
		if (args.length < 2) {
			System.err.println("Usage: jsva ExcelWriter templateFilePath outputFilePath");
			System.exit(1); // エラーの場合プログラムを終了する
		}
		
		String templateFilePath = args[0]; // テンプレートファイル
		String outputFilePath = args[1];  // アウトプットファイル
		// String jsonFilePath = args[2]; // 受け取るJSON形式ファイル ※試運転の際コメントアウト※
		// String cellConfigFilePath = "cellConfig.json"; // JSON設定ファイルのパス
		
		// VB.NET側のresourceディレクトリにJSON設定ファイルを配置し相対パスで指定
		String cellConfigFilePath = "cellConfig.json";
		
		FileInputStream fis = null;
		Workbook workbook = null;
		
		try {
			fis = new FileInputStream(templateFilePath);
			workbook = new XSSFWorkbook(fis);
			
			Sheet sheet = workbook.getSheetAt(0); // 最初のシートを取得
			
			// 設定ファイルを読み込む
			ClassLoader classLoader = ExcelWriter.class.getClassLoader();
			InputStream inputStream = classLoader.getResourceAsStream(cellConfigFilePath);
			if (inputStream == null) {
				throw new IOException("設定ファイルが見つかりません：" + cellConfigFilePath);
			}
			String cellConfigData = new String(inputStream.readAllBytes(), StandardCharsets.UTF_8); // JSON設定ファイルをString型でcellConfigDataに格納
			JSONObject cellConfig = new JSONObject(cellConfigData); // cellConfigDataをJSONオブジェクトで取得してcellConfig変数に格納
			
			// データの書き込み処理を記述
			if (sheet != null) {
				// VB.NET側の標準入力からJSONデータを読み込む
				BufferedReader reader = new BufferedReader(new InputStreamReader(System.in, StandardCharsets.UTF_8));
				StringBuilder sb = new StringBuilder();
				String line;
				while ((line = reader.readLine()) != null) {
					sb.append(line);
				}
				String jsonData = sb.toString();
				//System.out.println("jsonData: " + jsonData); // デバッグ用ログ出力
				JSONObject jsonObject = new JSONObject(jsonData);
				writeData(workbook, sheet, jsonObject, cellConfig);
			}
			
			// 印刷範囲を設定 例: A1からH最終行まで
			//int lastRowNum = sheet.getLastRowNum();
			workbook.setPrintArea(0, "$A$1:$M$37");
			
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
	public static void writeData(Workbook workbook, Sheet sheet, JSONObject jsonObject, JSONObject cellConfig) {
		JSONArray items = jsonObject.getJSONArray("items"); //JsonArray型(配列)の変数 items に格納
		// 設定ファイル cellConfig からキー "items" に対応するオブジェクトを取得 itemsConfigオブジェクトはitems配列の各要素をExcelのどのセルに書き込むかの定義する設定情報を持つ
		JSONObject itemsConfig = cellConfig.getJSONObject("items");
		writeItems(sheet, items, itemsConfig); // items配列の各要素をExcelシート (sheet) に書き込むためのメソッド
		/*Iterator<String> keys = jsonObject.keys(); //keys() : jsonオブジェクトに含まれるすべてのキーのイテレータを返す　Iterator<String> : 文字列型の要素を順番に所得するためのイテレータの型
		while (keys.hasNext()) { // whileループを使用し、keysが持つすべてのキーを順番に処理
			String key = keys.next(); // 次のキーを取得するメソッド
			if (key.equals("items")) { // キーが "items" と等しいとき
				JSONArray items = jsonObject.getJSONArray("items"); //JsonArray型(配列)の変数 items に格納
				// 設定ファイル cellConfig からキー "items" に対応するオブジェクトを取得 itemsConfigオブジェクトはitems配列の各要素をExcelのどのセルに書き込むかの定義する設定情報を持つ
				JSONObject itemsConfig = cellConfig.getJSONObject("items");
				writeItems(sheet, items, itemsConfig); // items配列の各要素をExcelシート (sheet) に書き込むためのメソッド
			} else { // キーが "items" と等しくない時
				String cellAddress = key; //keyはJSONオブジェクトのセルアドレスを表す
				String value = jsonObject.getString(key); // JSONオブジェクト (json) からキー (key) に対応する文字列の型の値を取得し、value 変数に格納する
				writeCell(sheet, cellAddress, value); // Excleシート (sheet) の指定されたセル (cellAddress) に値 (value) を書き込むメソッド
			}
		}*/
	}
	// items配列の各要素をExcelシート (sheet) に書き込むためのメソッド
	private static void writeItems(Sheet sheet, JSONArray items, JSONObject itemsConfig) {
		int startRow = itemsConfig.getInt("startRow"); // cellConfigの開始行 (17行目である16) を取得 startRow は書き始めるExcelシートの行番号
		int page2StartRow = itemsConfig.getInt("page2StartRow"); // 2ページ目の開始行 設定ファイルから取得（40行目である39）
		int page3StartRow = itemsConfig.getInt("page3StartRow"); // 3ページ目の開始行 設定ファイルから取得（76行目である75）
		int firstPageItemsPerPage = itemsConfig.getInt("firstPageItemsPerPage"); // 1ページ目は21項目を表示
		int secondPagesItemsPerPage = itemsConfig.getInt("secondPagesItemsPerPage"); // 2ページ目は34項目を表示
		int thirdPagesItemsPerPage = itemsConfig.getInt("thirdPagesItemsPerPage"); // 3ページ目は33項目を表示
		int totalAmount = 0; // H列の合計値を格納する変数
		int itemsLength = items.length(); // JSONArrayの長さ
		int writtenItems = 0; // 
		// 1ページ目の書き込み
		//for (int i = 0; i < items.length(); i++) { // itemsというJASONArrayの長さだけループする
		for (int i = 0; i < Math.min(itemsLength, firstPageItemsPerPage); i++) {
			int rowNum = startRow + i; // 行番号
			JSONObject item = items.getJSONObject(i); // items配列から i 番目の要素を取得
			writeItemRow(sheet, item, itemsConfig, rowNum);
			// H列の値を合計
			try {
				totalAmount += Integer.parseInt(item.getString("amount"));
			} catch (NumberFormatException e) {
				// 数値に変換できない場合は無視
			}
			writtenItems++;
		}
		// 2ページ目の改ページを設定
		if (itemsLength > firstPageItemsPerPage) {
			sheet.setRowBreak(37); // 38行目の前に改ページ
			// 2ページ目の書き込み
			for (int i = 0; i < Math.min(itemsLength - firstPageItemsPerPage, secondPagesItemsPerPage); i++) {
				int rowNum = page2StartRow + i;
				JSONObject item = items.getJSONObject(firstPageItemsPerPage + i);
				writeItemRow(sheet, item, itemsConfig, rowNum);
				try {
					totalAmount += Integer.parseInt(item.getString("amount"));
				} catch (NumberFormatException e) {
					// 数値に変換できない場合は無視
				}
				writtenItems++;
			}
		}
		// 3ページ目の改ページ設定
		if (itemsLength > firstPageItemsPerPage + secondPagesItemsPerPage) {
			sheet.setRowBreak(73); // 74行目の前に改ページ
			// 3ページ目の書き込み
			for (int i = 0; i < Math.min(itemsLength - firstPageItemsPerPage - secondPagesItemsPerPage, thirdPagesItemsPerPage); i++) {
				int rowNum = page3StartRow + i;
				JSONObject item = items.getJSONObject(firstPageItemsPerPage + secondPagesItemsPerPage + i);
				writeItemRow(sheet, item, itemsConfig, rowNum);
				try {
					totalAmount += Integer.parseInt(item.getString("amount"));
				} catch (NumberFormatException e){
					// 数値に変換できない場合は無視
				}
				writtenItems++;
			}
		}
		// 合計行の書き込み
		if (writtenItems > 0) {
			int totalRow;
			if (writtenItems <= firstPageItemsPerPage) {
				totalRow = startRow + firstPageItemsPerPage;
			} else if (writtenItems <= firstPageItemsPerPage + secondPagesItemsPerPage) {
				totalRow = page2StartRow + secondPagesItemsPerPage;
			} else {
				totalRow = page3StartRow + thirdPagesItemsPerPage;
			}
			writeCell(sheet, "B" + totalRow, "合計");
			writeCell(sheet, "H" + totalRow, String.valueOf(totalAmount));
		}
		// 合計行も書き込まれた項目としてカウントする
		//writtenItems++;
	}
	private static void writeItemRow(Sheet sheet, JSONObject item, JSONObject itemsConfig, int rowNum) {
		Row row = sheet.getRow(rowNum); // Excelシートの startRow + i 行目のオブジェクトを取得し row に代入
		if (row == null) {
			row = sheet.createRow(rowNum); // 取得しようとした行が null の場合、新しい行を作成し row に代入
		}
		// writeCell はExcelシートの指定されたセルに値を書き込む関数 
		// itemsConfig.getString("name") 設定ファイルから"name"に対応する列名(B列)を取得
		// item.getString("name") 現在の"name"に対応する値を取得
		writeCell(sheet, itemsConfig.getString("name") + (rowNum + 1), item.getString("name"));
		writeCell(sheet, itemsConfig.getString("quantity") + (rowNum + 1), item.getString("quantity"));
		writeCell(sheet, itemsConfig.getString("unit") + (rowNum + 1), item.getString("unit"));
		writeCell(sheet, itemsConfig.getString("price") + (rowNum + 1), item.getString("price"));
		writeCell(sheet, itemsConfig.getString("amount") + (rowNum + 1), item.getString("amount"));
		writeCell(sheet, itemsConfig.getString("subNote") + (rowNum + 1), item.getString("subNote"));
	}
	public static void writeCell(Sheet sheet, String cellAddress, String cellvalue) { // cellAddress 書き込み先のセルアドレス ("A1", "B5")を表す文字列
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
	}
}