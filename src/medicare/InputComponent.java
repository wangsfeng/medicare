package medicare;

import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.event.ActionListener;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

public class InputComponent {
	public String file = "";// 导出excel路径
//	public static String beginTime = "";
//	public static String endTime = "";
	public int reportNo = 0;// 报表号 默认0号报表

	JFrame f = new JFrame("医保数据----山西省肿瘤医院");
	// 菜单栏
	JMenuBar mb = new JMenuBar();
	JMenu menu = new JMenu("医保数据");
	private CardLayout card = new CardLayout(5, 5);
	JPanel cardPanel = new JPanel(card);

	JFileChooser fc1 = new JFileChooser();

	public void init() {
		// 新建报表panel,然后加载到cardPanel
		ReportPanel[] p = new ReportPanel[ReportNameConstants.report.length];
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			p[i] = new ReportPanel(ReportNameConstants.report[i]);
			cardPanel.add(p[i], String.valueOf(i));// 加入card中
		}

		// 菜单监听器
		ActionListener al = e -> {
			for (int i = 0; i < ReportNameConstants.report.length; i++) {
				if (e.getActionCommand() == ReportNameConstants.report[i]) {
					card.show(cardPanel, String.valueOf(i));
					reportNo = i;
					break;
				}
			}
		};
		// 文件选择器监听器
		ActionListener fileChooser = e -> {
			fc1.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);// 只能选路径,不能选文件
			fc1.showOpenDialog(f);// 显示JFileChooser
			fc1.setFileSelectionMode(1);// 设定只能选择到文件夹
			file = fc1.getSelectedFile().getAbsolutePath();
		};
		// 开始导出button监听器
		ActionListener beginOut = e -> {
			// 获取开始时间、结束时间
			Object x = e.getSource();// 取得了调用监听器的button
			for (int i = 0; i < ReportNameConstants.report.length; i++) {
				if (x.equals(p[i].getComponent(6))) {
					ReportTimeConstants.beginTime = ((DateChooserJButton) p[i].getComponent(2)).getText();
					ReportTimeConstants.endTime = ((DateChooserJButton) p[i].getComponent(4)).getText();
				}
			}
			new OutputExcel(ReportTimeConstants.beginTime, ReportTimeConstants.endTime, file, reportNo);// 调用导出类
		};

		// f.setBounds(200, 200,500,500);
		// 加载菜单
		JMenuItem[] mi = new JMenuItem[ReportNameConstants.report.length];
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			mi[i] = new JMenuItem(ReportNameConstants.report[i]);
			mi[i].addActionListener(al);
			menu.add(mi[i]);
		}
		mb.add(menu);
		f.add(mb, BorderLayout.NORTH);
		// 加载cardPanel
		f.add(cardPanel);
		for (int i = 0; i < ReportNameConstants.report.length; i++) {
			p[i].out.addActionListener(fileChooser);// 加载文件选择器监听器
			p[i].bn.addActionListener(beginOut);// 加载开始导出button监听器
		}
		// 设置外观风格
		try {
			// javax.swing.plaf.nimbus.NimbusLookAndFeel
			// com.sun.java.swing.plaf.windows.WindowsLookAndFeel
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
			SwingUtilities.updateComponentTreeUI(f);
			SwingUtilities.updateComponentTreeUI(fc1);
		} catch (ClassNotFoundException | InstantiationException | IllegalAccessException
				| UnsupportedLookAndFeelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		f.pack();
		// 显示总体窗口
		f.setVisible(true);

	}

	public static void main(String[] args) {
		new InputComponent().init();
	}
}
