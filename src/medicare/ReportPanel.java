package medicare;

import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JPanel;

public class ReportPanel extends JPanel {

	// 数据报表输入界面
	public JLabel begin = new JLabel("开始时间");
	public DateChooserJButton beginTime = new DateChooserJButton(ReportTimeConstants.beginTime);
	public JLabel end = new JLabel("结束时间");
	public DateChooserJButton endTime = new DateChooserJButton(ReportTimeConstants.endTime);
	public JButton out = new JButton("选择导出路径");
	public JButton bn = new JButton("开始导出");

	// 构造函数
	public ReportPanel(String arg) {
		super();
		JLabel name = new JLabel(arg);
		// 加载数据报表输入界面
		this.add(name);
		this.add(begin);
		this.add(beginTime);
		this.add(end);
		this.add(endTime);
		this.add(out);
		this.add(bn);
	}
}
