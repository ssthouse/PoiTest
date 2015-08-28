package poi;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.poi.POIXMLDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;

public class AnotherTest {

	public static void searchAndReplace(String srcPath, String destPath,
			Map<String, String> map) {
		try {
			XWPFDocument document = new XWPFDocument(
					POIXMLDocument.openPackage(srcPath));
			// 替换段落中的指定文字
			Iterator<XWPFParagraph> itPara = document.getParagraphsIterator();
			while (itPara.hasNext()) {
				XWPFParagraph paragraph = (XWPFParagraph) itPara.next();
				// String s = paragraph.getParagraphText();
				Set<String> set = map.keySet();
				Iterator<String> iterator = set.iterator();
				while (iterator.hasNext()) {
					String key = iterator.next();
					List<XWPFRun> run = paragraph.getRuns();
					for (int i = 0; i < run.size(); i++) {
						if (run.get(i).getText(run.get(i).getTextPosition()) != null
								&& run.get(i)
										.getText(run.get(i).getTextPosition())
										.equals(key)) {
							/**
							 * 参数0表示生成的文字是要从哪一个地方开始放置,设置文字从位置0开始 就可以把原来的文字全部替换掉了
							 * */
							run.get(i).setText(map.get(key), 0);
						}
					}
				}
			}

			// 替换表格中的指定文字
			Iterator<XWPFTable> itTable = document.getTablesIterator();
			while (itTable.hasNext()) {
				XWPFTable table = (XWPFTable) itTable.next();
				int rcount = table.getNumberOfRows();
				for (int i = 0; i < rcount; i++) {
					XWPFTableRow row = table.getRow(i);
					List<XWPFTableCell> cells = row.getTableCells();
					for (XWPFTableCell cell : cells) {
						for (Entry<String, String> e : map.entrySet()) {
							if (cell.getText().equals(e.getKey())) {
								cell.removeParagraph(0);
								cell.setText(e.getValue());
							}
						}
					}
				}
			}
			FileOutputStream outStream = null;
			outStream = new FileOutputStream(destPath);
			document.write(outStream);
			outStream.close();
		} catch (Exception e) {
			e.printStackTrace();
		}

	}

	public static void main(String[] args) throws Exception {
		Map<String, String> map = new HashMap<String, String>();
		map.put("${projectName}", "王五王五王五王五王五王五王五王五王五王五王五王五");
		map.put("${temperatureItem1}", "woquearjgakrejgbarekj");
		String srcPath = "F:\\test.docx";
		String destPath = "F:\\output.docx";
		searchAndReplace(srcPath, destPath, map);
	}
}
