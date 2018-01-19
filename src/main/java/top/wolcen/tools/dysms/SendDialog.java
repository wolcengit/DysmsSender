package top.wolcen.tools.dysms;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

import javax.swing.JButton;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JTextField;
import javax.swing.SwingWorker;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileFilter;

import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.fife.ui.rsyntaxtextarea.SyntaxConstants;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import com.aliyuncs.exceptions.ClientException;

import jxl.Sheet;
import jxl.Workbook;

public class SendDialog extends JDialog {
	private static Logger logger = Logger.getLogger(SendDialog.class);

	private boolean running = false;
	
	private final JPanel contentPanel = new JPanel();
	private JTextField tfFile;
	private JButton btnFile;
	private JLabel lblConf;
	private SyntaxEditor seConf;
	private JButton okButton;
	private JButton cancelButton;
	private JButton btnLoad;
	private JButton btnSave;
	private JLabel lblMsg;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		try {
			System.setProperty("file.encoding", "UTF-8");

			// 默认使用系统代理
			System.setProperty("java.net.useSystemProxies", "true");
			
			String configFilename = System.getProperty("user.dir") + "/log4j.properties";
			File f = new File(configFilename);
			if(!f.exists()){
				try {
					PrintWriter pw = new PrintWriter(new OutputStreamWriter(new FileOutputStream(f)), true);
					pw.println("#rootLogger");
					pw.println("log4j.rootLogger=debug,stdout,file");
					pw.println("#console");
					pw.println("log4j.appender.stdout=org.apache.log4j.ConsoleAppender");
					pw.println("log4j.appender.stdout.Target=System.out");
					pw.println("log4j.appender.stdout.layout=org.apache.log4j.PatternLayout");
					pw.println("log4j.appender.stdout.layout.ConversionPattern=%d [%5p] %l - %m%n");
					pw.println("#file");
					pw.println("log4j.appender.file=org.apache.log4j.DailyRollingFileAppender");
					pw.println("log4j.appender.file.File=${user.dir}//logs//error.log");
					pw.println("log4j.appender.file.DatePattern = '.'yyyy-MM-dd ");
					pw.println("log4j.appender.file.layout=org.apache.log4j.PatternLayout");
					pw.println("log4j.appender.file.layout.ConversionPattern=%d [%5p] %l - %m%n");
					pw.flush();
					pw.close();
				} catch (Exception e) {
				}
			}

			PropertyConfigurator.configure(configFilename);
			
			SendDialog dialog = new SendDialog();
			dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
			dialog.setVisible(true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * Create the dialog.
	 */
	public SendDialog() {
		initComponents();
		
		this.seConf.setText(AliyunSmsSender.confTemplate());
		
		this.lblMsg = new JLabel("");
		this.lblMsg.setForeground(Color.MAGENTA);
		this.lblMsg.setBounds(52, 361, 552, 23);
		contentPanel.add(this.lblMsg);
		
	}
	private void initComponents() {
		setTitle("阿里云短信发送工具");
		setBounds(100, 100, 637, 473);
		getContentPane().setLayout(new BorderLayout());
		this.contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));
		getContentPane().add(this.contentPanel, BorderLayout.CENTER);
		contentPanel.setLayout(null);
		{
			JLabel lblFile = new JLabel("名单:");
			lblFile.setBounds(10, 10, 51, 21);
			contentPanel.add(lblFile);
		}
		
		this.tfFile = new JTextField();
		this.tfFile.setBounds(52, 10, 485, 21);
		contentPanel.add(this.tfFile);
		this.tfFile.setColumns(10);
		
		this.btnFile = new JButton("...");
		this.btnFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				do_btnFile_actionPerformed(e);
			}
		});
		this.btnFile.setBounds(553, 9, 51, 23);
		contentPanel.add(this.btnFile);
		
		this.lblConf = new JLabel("配置:");
		this.lblConf.setBounds(10, 41, 51, 15);
		contentPanel.add(this.lblConf);
		
		this.seConf = new SyntaxEditor();
		this.seConf.setEditable(true);
		this.seConf.setBounds(52, 41, 552, 310);
		this.seConf.setSyntaxStyle(SyntaxConstants.SYNTAX_STYLE_JSON);
		contentPanel.add(this.seConf);
		{
			JPanel buttonPane = new JPanel();
			buttonPane.setPreferredSize(new Dimension(10, 40));
			getContentPane().add(buttonPane, BorderLayout.SOUTH);
			buttonPane.setLayout(null);
			{
				okButton = new JButton("发送短信");
				okButton.addActionListener(new ActionListener() {
					public void actionPerformed(ActionEvent e) {
						do_okButton_actionPerformed(e);
					}
				});
				okButton.setBounds(358, 10, 112, 25);
				okButton.setActionCommand("OK");
				buttonPane.add(okButton);
				getRootPane().setDefaultButton(okButton);
			}
			{
				cancelButton = new JButton("退出系统");
				cancelButton.addActionListener(new ActionListener() {
					public void actionPerformed(ActionEvent e) {
						do_cancelButton_actionPerformed(e);
					}
				});
				cancelButton.setBounds(486, 10, 112, 25);
				cancelButton.setActionCommand("Cancel");
				buttonPane.add(cancelButton);
			}
			
			btnLoad = new JButton("载入配置");
			btnLoad.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					do_btnLoad_actionPerformed(e);
				}
			});
			btnLoad.setBounds(10, 11, 112, 23);
			buttonPane.add(btnLoad);
			
			btnSave = new JButton("保存配置");
			btnSave.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent e) {
					do_btnSave_actionPerformed(e);
				}
			});
			btnSave.setBounds(132, 11, 112, 23);
			buttonPane.add(btnSave);
		}
	}
	protected void do_btnFile_actionPerformed(ActionEvent e) {
		JFileChooser jfc = new JFileChooser();
		jfc.setCurrentDirectory(new File("."));
		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "名单Excel File";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory())
					return true;
				String name = f.getName();
				return name.endsWith(".xls");
			}
		});
		int r = jfc.showDialog(this, btnFile.getText());
		if (r == JFileChooser.APPROVE_OPTION) {
			File file = jfc.getSelectedFile();
			tfFile.setText(file.getAbsolutePath());
		}
		
	}
	protected void do_btnLoad_actionPerformed(ActionEvent e) {
		JFileChooser jfc = new JFileChooser();
		jfc.setCurrentDirectory(new File("."));
		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "配置文件";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory())
					return true;
				String name = f.getName();
				return name.endsWith(".jcf");
			}
		});
		int r = jfc.showDialog(this, btnFile.getText());
		if (r == JFileChooser.APPROVE_OPTION) {
			File file = jfc.getSelectedFile();
			
			try {
				String string = IOUtils.toString(new FileInputStream(file));
				seConf.setText(string);
			} catch (FileNotFoundException e1) {
				e1.printStackTrace();
				logger.error("ERROR load", e1);
			} catch (IOException e1) {
				e1.printStackTrace();
				logger.error("ERROR load", e1);
			}
			
		}
	}
	protected void do_btnSave_actionPerformed(ActionEvent e) {
		JFileChooser jfc = new JFileChooser();
		jfc.setCurrentDirectory(new File("."));
		jfc.setFileSelectionMode(JFileChooser.FILES_ONLY);
		jfc.setFileFilter(new FileFilter() {

			@Override
			public String getDescription() {
				return "配置文件";
			}

			@Override
			public boolean accept(File f) {
				if (f.isDirectory())
					return true;
				String name = f.getName();
				return name.endsWith(".jcf");
			}
		});
		int r =jfc.showSaveDialog(btnSave);
		if (r == JFileChooser.APPROVE_OPTION) {
			File file = jfc.getSelectedFile();
			if(!file.getName().endsWith(".jcf")){
				file = new File(file.getAbsolutePath()+".jcf");
			}
			String string = seConf.getText();
			try {
				IOUtils.write(string, new FileOutputStream(file));
			} catch (IOException e1) {
				e1.printStackTrace();
				logger.error("ERROR save", e1);
			}
		}
	}
	protected void do_okButton_actionPerformed(ActionEvent e) {
		String fileName = tfFile.getText();
		String string = seConf.getText();
		if(StringUtils.isBlank(fileName)){
			JOptionPane.showMessageDialog(okButton, "请选择名单Excel先");
			return;
		}
		JSONObject jso = null;
		try {
			jso = JSON.parseObject(string);
		} catch (Exception e1) {
			e1.printStackTrace();
			logger.error("ERROR json format", e1);
			JOptionPane.showMessageDialog(okButton, "请输入配置先（格式错误）");
			return;
		}
		if(StringUtils.isBlank(string) || jso == null){
			JOptionPane.showMessageDialog(okButton, "请输入配置先");
			return;
		}
		if(!jso.containsKey("accessKey") ||!jso.containsKey("accessSecret") ||!jso.containsKey("signCode") || !jso.containsKey("templateCode") ||!jso.containsKey("rowStart") ||!jso.containsKey("rowEnd") ||!jso.containsKey("recNum")){
			JOptionPane.showMessageDialog(okButton, "请输入正确的配置先，可能缺少【accessKey,accessSecret,signCode,templateCode,rowStart,rowEnd,recNum】");
			return;
		}
		
		running = true;
		enabledComponent(false);
		new SwingWorker<Void, Void>() {

			@Override
			protected Void doInBackground() throws Exception {
				showMessage("开始发送.....");
				sendProcess();
				return null;
			}

			@Override
			protected void done() {
				running = false;
				enabledComponent(true);
				super.done();
			}
			
			
			
		}.execute();
		
		
	}
	protected void do_cancelButton_actionPerformed(ActionEvent e) {
		if(running){
			JOptionPane.showMessageDialog(cancelButton, "正在运行.....");
			return;
		}
		System.exit(0);
		
	}
	
	public void enabledComponent(boolean enabled){
		this.tfFile.setEnabled(enabled);
		this.btnFile.setEnabled(enabled);
		this.seConf.setEnabled(enabled);
		this.btnLoad.setEnabled(enabled);
		this.btnSave.setEnabled(enabled);
		this.okButton.setEnabled(enabled);
		this.cancelButton.setEnabled(enabled);
	}
	
	public void showMessage(String text) {
		this.lblMsg.setText(text);
		this.lblMsg.invalidate();
	}
	
	public int strToColumn(String str) {
		char[] ch = str.toCharArray();
		if (ch.length == 1) {
			return ((int) ch[0]) - 65;
		} else {
			return 26 * (((int) ch[0]) - 64) + (((int) ch[1]) - 65);
		}
	}
	
	public void sendProcess(){
		String fileName = tfFile.getText();
		String string = seConf.getText();
		JSONObject jso = JSON.parseObject(string);

		String accessKey = jso.getString("accessKey");
		String accessSecret = jso.getString("accessSecret");
		String signCode = jso.getString("signCode");
		String templateCode = jso.getString("templateCode");
		String recNum = jso.getString("recNum");
		
		int rowStart = jso.getIntValue("rowStart") - 1;
		int rowEnd = jso.getIntValue("rowEnd") - 1;
		
		logger.info("signCode:"+signCode+"  templateCode:"+templateCode+"  rowStart:"+(rowStart+1)+"  rowEnd:"+(rowEnd+1));

		boolean over = false;
		try {
			Workbook workbook = null;
			workbook = Workbook.getWorkbook(new File(fileName));
			Sheet sheet = workbook.getSheet(0);
			if(rowEnd > sheet.getRows() ) rowEnd = sheet.getRows();
			for (int r = rowStart; r <= rowEnd; r++) {
				
				JSONObject jsp = new JSONObject();
				for(String key:jso.keySet()){
					if("accessKey".equals(key)) continue;
					if("accessSecret".equals(key)) continue;
					if("signCode".equals(key)) continue;
					if("templateCode".equals(key)) continue;
					if("rowStart".equals(key)) continue;
					if("rowEnd".equals(key)) continue;
					if("recNum".equals(key)) continue;
					String val = jso.getString(key);
					if(val.startsWith("@")){
						jsp.put(key, val.substring(1));
					}else{
						String value = sheet.getCell(strToColumn(val), r).getContents().trim();
						jsp.put(key, value);
					}
				}
				String params = jsp.toJSONString(jsp);
				String recnum = sheet.getCell(strToColumn(recNum), r).getContents().trim();
				
				showMessage("开始发送:"+recnum);
				logger.info("recNum:"+recnum+"  Params:"+params);
				
				String rslt;
				try {
					rslt = AliyunSmsSender.send(accessKey, accessSecret, signCode, templateCode, params, recnum);
					logger.info("Reslt:"+rslt);
				} catch (ClientException e) {
					logger.error("发送失败", e);
					e.printStackTrace();
					over = true;
					showMessage("发送失败:"+e.getErrMsg());
				}
			   if(over) break;
			}
			workbook.close();	
			if(!over){
				showMessage("发送完成");
			}
			
		} catch (Exception e) {
			logger.error("发送处理失败", e);
			e.printStackTrace();
			showMessage("发送处理失败:"+e.getMessage());
		}
		
	}
	
}
