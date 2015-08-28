package poi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.sql.PseudoColumnUsage;
import java.util.HashMap;
import java.util.Map;

import jframw.Log;

import org.apache.poi.POIDocument;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.util.POIUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordHelper {

	//操作的文件路径
	private static String SRC_PATH;
	static {
		SRC_PATH = System.getProperty("user.dir") + "/src/";
	}
	
	private static WordHelper wordHelper;
	
	private WordHelper(){}
	
	public static WordHelper getInstance(){
		if(wordHelper == null){
			wordHelper = new WordHelper();
		}
		return wordHelper;
	}



	/**
	 * 生成word文档的方法
	 *
	 * @param context
	 * @param tourItem
	 */
	public void generateWordFile(Map<String , String> mapData) {
		try {
			// 读取word模板
			// TODO--打开模板文件的输入流
			InputStream is = getModelWord();

			HWPFDocument doc = new HWPFDocument(is);
			// 读取word文本操作模块
			Range range = doc.getRange();
			// 替换文本内容
			for (Map.Entry<String, String> oneData : mapData.entrySet()) {
				range.replaceText(oneData.getKey(), oneData.getValue());
			}
			// 以工程名+"-"+touNumebr作为word的文件名
			String fileName = "newWordFile.doc";
			FileOutputStream out = new FileOutputStream(SRC_PATH + fileName);
			doc.write(out);
			// 处理流
			out.flush();
			out.close();
			is.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 获取Model文件输入流
	 * @return
	 */
	public InputStream getModelWord() {
		InputStream is = null;
		try {
			is = new FileInputStream(SRC_PATH+"/model.doc");
			Log.logError(SRC_PATH+"/model.doc");
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			Log.logError("file not found");
			e.printStackTrace();
		}
		return is;
	}
}
