package jframw;

import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.util.HashMap;
import java.util.Map;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFormattedTextField;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.UIManager;
import javax.swing.border.EmptyBorder;

import poi.WordHelper;
import poi.WordKey;
import poi.XwpfTest;

public class Frame1 extends JFrame {

	private JPanel contentPane;
	private JButton btnGenerateWord;
	private JTextField textFieldItem2;
	private JFormattedTextField textFieldTitle;
	private JFormattedTextField textFieldItem1;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		try {
			Frame1 frame = new Frame1();
			frame.setVisible(true);
		} catch (Exception e) {
			e.printStackTrace();
		}

		try {
			new XwpfTest().testTemplateWrite();
//			Log.logError("替换完成");
		} catch (Exception e) {
			Log.logError("出错了");
			e.printStackTrace();
		}
	}

	/**
	 * 获取LookAndFeel数据
	 */
	private static void showLAF() {
		for (UIManager.LookAndFeelInfo lookAndFeelInfo : UIManager
				.getInstalledLookAndFeels()) {
			Log.logError(lookAndFeelInfo.getName());
		}
	}

	private void generateReplaceMap() {

	}

	/**
	 * 获取数据的映射表
	 *
	 * @param tourItem
	 * @return
	 */
	private Map<String, String> getDataMap() {
		// 获取各个Item的数据
		String title = textFieldTitle.getText();
		String item1 = textFieldItem1.getText();
		String item2 = textFieldItem2.getText();
		Log.logError(title + "    " + item1 + "   " + item2);
		// 填充map数据
		Map<String, String> dataMap = new HashMap<>();
		dataMap.put(WordKey.PRJ_NAME, title);
		return dataMap;
	}

	/**
	 * Create the frame.
	 */
	public Frame1() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		contentPane.setLayout(new BorderLayout(0, 0));
		setContentPane(contentPane);

		JPanel panelUp = new JPanel();
		contentPane.add(panelUp, BorderLayout.CENTER);
		panelUp.setLayout(new BorderLayout(0, 0));

		JPanel panel = new JPanel();
		panelUp.add(panel, BorderLayout.CENTER);
		panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));

		JLabel lbltitle = new JLabel("请输入Title");
		panel.add(lbltitle);

		textFieldTitle = new JFormattedTextField();
		panel.add(textFieldTitle);

		JLabel lblitem = new JLabel("请输入Item1");
		panel.add(lblitem);

		textFieldItem1 = new JFormattedTextField();
		panel.add(textFieldItem1);

		JLabel lblitem_1 = new JLabel("请输入Item2");
		panel.add(lblitem_1);

		textFieldItem2 = new JTextField();
		panel.add(textFieldItem2);
		textFieldItem2.setColumns(10);

		JPanel panel_1 = new JPanel();
		panelUp.add(panel_1, BorderLayout.LINE_END);
		panel_1.setLayout(new BoxLayout(panel_1, BoxLayout.Y_AXIS));

		JPanel panelBottom = new JPanel();
		contentPane.add(panelBottom, BorderLayout.SOUTH);

		btnGenerateWord = new JButton("生成Word文档");
		btnGenerateWord.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent e) {
				// TODO---生成word文档
				WordHelper.getInstance().generateWordFile(getDataMap());
			}
		});
		panelBottom.add(btnGenerateWord);
	}
}
