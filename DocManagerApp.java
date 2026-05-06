package org.example;

// 导入Word旧版本文档解析工具类
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
// 导入Word新版本文档解析工具类
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
// 导入PDF文档解析工具类
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
// 导入Java界面可视化组件
import javax.swing.*;
// 导入Java界面布局、颜色工具类
import java.awt.*;
// 导入文件读取、写入的流工具类
import java.io.*;
// 导入MySQL数据库操作工具类
import java.sql.*;
// 导入集合工具类，用于存储解析出来的标题数据
import java.util.ArrayList;
import java.util.List;

/**
 * 项目名称：技术文档管理与结构化解析系统
 * 开发环境：Java8 + IDEA + MySQL8.0
 * 核心功能：支持doc/docx/pdf/txt文档上传、标题解析、目录生成、结构检查、文档搜索
 * 用途：软件工程基础实训答辩项目
 */
public class DocManagerApp {
    // ========================= 数据库配置板块 =========================
    // 作用：固定配置MySQL数据库的连接信息，所有数据库操作都依赖这部分代码
    // 本地数据库连接地址，doc_manager是我们创建的数据库名
    private static final String URL = "jdbc:mysql://localhost:3306/doc_manager?useSSL=false&serverTimezone=UTC";
    // MySQL默认用户名
    private static final String USER = "root";
    // 个人MySQL登录密码（答辩时说明这里是自己的数据库密码）
    private static final String PASSWORD = "heping520";

    // ========================= 数据库连接工具方法 =========================
    // 功能：获取数据库的连接对象，是程序和数据库沟通的桥梁
    // 实现：加载数据库驱动，建立连接，连接失败弹出提示框
    public static Connection getConnection() {
        try {
            // 加载MySQL官方驱动类
            Class.forName("com.mysql.cj.jdbc.Driver");
            // 根据配置信息，创建并返回数据库连接
            return DriverManager.getConnection(URL, USER, PASSWORD);
        } catch (Exception e) {
            // 连接失败，弹出错误提示
            JOptionPane.showMessageDialog(null, "数据库连接失败！");
            return null;
        }
    }

    // ========================= 标题层级判断工具方法 =========================
    // 功能：识别文档中的一行文字是 一级标题/二级标题/普通文本
    // 规则：第X章=一级标题；1.1/1=二级标题；其他文字=普通文本
    // 返回值：1=一级标题 2=二级标题 0=普通文本
    private static int getTitleLevel(String text) {
        // 匹配一级标题：以"第"开头，并且包含"章"（例如：第一章 项目概述）
        if (text.startsWith("第") && text.contains("章")) {
            return 1;
        }
        // 匹配二级标题：数字+点+数字格式（例如：1.1 项目背景）
        if (text.matches("^\\d+\\.\\d+.*")) {
            return 2;
        }
        // 匹配二级标题：纯数字开头（例如：1 项目介绍）
        if (text.matches("^\\d+.*")) {
            return 2;
        }
        // 不属于标题，返回0
        return 0;
    }

    // ========================= 核心功能：文档上传与解析 =========================
    // 功能：用户选择文件，程序自动识别格式并解析标题，最终存入数据库
    // 支持格式：.doc(旧Word)、.docx(新Word)、.pdf、.txt
    public static void uploadFile() {
        // 创建文件选择窗口，让用户选择本地文档
        JFileChooser chooser = new JFileChooser();
        // 设置文件过滤器：只显示我们支持的4种文档格式
        chooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                String name = f.getName().toLowerCase();
                // 允许选择文件夹，以及4种格式的文档
                return f.isDirectory() || name.endsWith(".doc") || name.endsWith(".docx")
                        || name.endsWith(".pdf") || name.endsWith(".txt");
            }
            @Override
            public String getDescription() {
                return "文档文件(*.doc;*.docx;*.pdf;*.txt)";
            }
        });

        // 打开文件选择框，用户未选择文件则直接退出方法
        int res = chooser.showOpenDialog(null);
        if (res != JFileChooser.APPROVE_OPTION) {
            return;
        }

        // 获取用户选择的文件、文件路径、文件名
        File file = chooser.getSelectedFile();
        String path = file.getAbsolutePath();
        String name = file.getName();
        // 创建集合，用于存储解析出来的所有标题信息
        List<ParagraphInfo> list = new ArrayList<>();

        try {
            // ========== 1. 解析TXT文本文件 ==========
            if (name.endsWith(".txt")) {
                // 以UTF-8编码读取TXT文件，防止中文乱码
                BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
                String line;
                // 逐行读取文件内容
                while ((line = br.readLine()) != null) {
                    line = line.trim();
                    // 判断当前行是否为标题，是标题就加入集合
                    int level = getTitleLevel(line);
                    if (level > 0) {
                        list.add(new ParagraphInfo(level, line));
                    }
                }
                br.close();
            }

            // ========== 2. 解析PDF文件 ==========
            else if (name.endsWith(".pdf")) {
                // 加载PDF文档
                PDDocument doc = PDDocument.load(file);
                // 提取PDF中的所有文本内容
                PDFTextStripper stripper = new PDFTextStripper();
                String text = stripper.getText(doc);
                doc.close();
                // 按行分割文本
                String[] lines = text.split("\\r?\\n");
                // 遍历每一行，判断是否为标题
                for (String line : lines) {
                    line = line.trim();
                    int level = getTitleLevel(line);
                    if (level > 0) {
                        list.add(new ParagraphInfo(level, line));
                    }
                }
            }

            // ========== 3. 解析DOCX新版Word文件 ==========
            else if (name.endsWith(".docx")) {
                XWPFDocument doc = new XWPFDocument(new FileInputStream(file));
                // 遍历Word中的所有段落
                for (XWPFParagraph p : doc.getParagraphs()) {
                    String line = p.getText().trim();
                    int level = getTitleLevel(line);
                    if (level > 0) {
                        list.add(new ParagraphInfo(level, line));
                    }
                }
                doc.close();
            }

            // ========== 4. 解析DOC旧版Word文件 ==========
            else if (name.endsWith(".doc")) {
                HWPFDocument doc = new HWPFDocument(new FileInputStream(file));
                Range range = doc.getRange();
                // 遍历文档所有段落
                for (int i = 0; i < range.numParagraphs(); i++) {
                    Paragraph p = range.getParagraph(i);
                    String line = p.text().trim();
                    int level = getTitleLevel(line);
                    if (level > 0) {
                        list.add(new ParagraphInfo(level, line));
                    }
                }
                doc.close();
            }

            // ========== 将解析完成的标题数据存入数据库 ==========
            Connection conn = getConnection();
            if (conn == null) {
                return;
            }

            // 第一条SQL：向document表插入 文件名、文件路径
            String sql1 = "INSERT INTO document(filename,filepath) VALUES(?,?)";
            PreparedStatement ps1 = conn.prepareStatement(sql1, Statement.RETURN_GENERATED_KEYS);
            ps1.setString(1, name);
            ps1.setString(2, path);
            ps1.executeUpdate();
            // 获取插入后自动生成的文档ID
            ResultSet rs = ps1.getGeneratedKeys();
            int docId = 0;
            if (rs.next()) {
                docId = rs.getInt(1);
            }

            // 第二条SQL：向doc_section表插入 标题层级、标题内容
            String sql2 = "INSERT INTO doc_section(doc_id,level,content) VALUES(?,?,?)";
            PreparedStatement ps2 = conn.prepareStatement(sql2);
            // 批量插入所有标题数据
            for (ParagraphInfo info : list) {
                ps2.setInt(1, docId);
                ps2.setInt(2, info.level);
                ps2.setString(3, info.content);
                ps2.addBatch();
            }
            ps2.executeBatch();
            conn.close();
            // 解析保存成功，弹出提示
            JOptionPane.showMessageDialog(null, "✅ 解析并保存成功！");
        } catch (Exception e) {
            // 解析失败，弹出错误提示
            JOptionPane.showMessageDialog(null, "❌ 解析失败：" + e.getMessage());
            e.printStackTrace();
        }
    }

    // ========================= 功能模块：自动生成目录 =========================
    // 功能：从数据库读取最新上传文档的标题，按层级生成规范目录
    // 实现：二级标题自动缩进，弹窗展示目录
    public static void buildToc() {
        try {
            Connection conn = getConnection();
            if (conn == null) {
                return;
            }
            // 查询数据库中【最新上传】的文档ID
            Statement stmt = conn.createStatement();
            ResultSet rsDoc = stmt.executeQuery("SELECT id FROM document ORDER BY upload_time DESC LIMIT 1");
            if (!rsDoc.next()) {
                JOptionPane.showMessageDialog(null, "暂无文档");
                return;
            }
            int docId = rsDoc.getInt(1);

            // 根据文档ID，查询所有标题信息
            PreparedStatement ps = conn.prepareStatement("SELECT level,content FROM doc_section WHERE doc_id=? ORDER BY id");
            ps.setInt(1, docId);
            ResultSet rs = ps.executeQuery();

            // 拼接目录文本：二级标题缩进2个空格
            StringBuilder sb = new StringBuilder();
            while (rs.next()) {
                int level = rs.getInt("level");
                String content = rs.getString("content");
                // 一级标题不缩进，二级标题缩进
                for (int i = 0; i < level - 1; i++) {
                    sb.append("  ");
                }
                sb.append(content).append("\n");
            }

            // 创建新窗口，展示生成的目录
            JFrame f = new JFrame("目录");
            f.setSize(600, 400);
            JTextArea ta = new JTextArea(sb.toString());
            f.add(new JScrollPane(ta));
            f.setVisible(true);
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ========================= 功能模块：文档结构规范化检查 =========================
    // 功能：检查标题是否跳级（例如：一级标题直接跳到三级标题）
    // 实现：遍历标题层级，判断相邻标题差值是否大于1
    public static void checkStruct() {
        try {
            Connection conn = getConnection();
            if (conn == null) {
                return;
            }
            // 获取最新上传的文档ID
            ResultSet rsDoc = conn.createStatement().executeQuery("SELECT id FROM document ORDER BY upload_time DESC LIMIT 1");
            if (!rsDoc.next()) {
                JOptionPane.showMessageDialog(null, "暂无文档");
                return;
            }
            int docId = rsDoc.getInt(1);

            // 查询文档的所有标题
            PreparedStatement ps = conn.prepareStatement("SELECT level,content FROM doc_section WHERE doc_id=? ORDER BY id");
            ps.setInt(1, docId);
            ResultSet rs = ps.executeQuery();

            // 存储结构错误信息
            List<String> err = new ArrayList<>();
            int last = 0;
            // 遍历标题，判断是否跳级
            while (rs.next()) {
                int level = rs.getInt("level");
                String content = rs.getString("content");
                // 当前标题层级 - 上一个标题层级 > 1，说明跳级，记录错误
                if (level - last > 1) {
                    err.add("层级错误：" + content);
                }
                last = level;
            }

            // 弹出检查结果
            if (err.isEmpty()) {
                JOptionPane.showMessageDialog(null, "✅ 结构规范");
            } else {
                JOptionPane.showMessageDialog(null, String.join("\n", err));
            }
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ========================= 功能模块：文档关键词搜索 =========================
    // 功能：根据用户输入的关键词，搜索已上传的文档（模糊匹配文件名）
    public static void search(String key) {
        try {
            Connection conn = getConnection();
            if (conn == null) {
                return;
            }
            // 模糊查询：文件名包含关键词就匹配
            PreparedStatement ps = conn.prepareStatement("SELECT filename FROM document WHERE filename LIKE ?");
            ps.setString(1, "%" + key + "%");
            ResultSet rs = ps.executeQuery();

            // 封装查询结果
            List<String> list = new ArrayList<>();
            while (rs.next()) {
                list.add(rs.getString(1));
            }

            // 弹出搜索结果
            if (list.isEmpty()) {
                JOptionPane.showMessageDialog(null, "无结果");
            } else {
                JOptionPane.showMessageDialog(null, String.join("\n", list));
            }
            conn.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // ========================= 程序入口：主界面（美化版） =========================
    // 功能：创建可视化操作界面，绑定所有功能按钮
    // 优化：窗口居中、布局整齐、字体统一、界面美观
    public static void main(String[] args) {
        // 1. 主窗口基础设置
        JFrame f = new JFrame("文档管理系统");
        f.setSize(500, 350);
        f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        f.setLocationRelativeTo(null); // 窗口居中显示
        f.setLayout(new BorderLayout(10, 10)); // 边界布局，带间距
        f.getContentPane().setBackground(new Color(245, 245, 245)); // 浅灰色背景

        // 2. 统一设置字体，保证界面美观
        Font titleFont = new Font("微软雅黑", Font.BOLD, 20);
        Font btnFont = new Font("微软雅黑", Font.PLAIN, 14);
        Font tipFont = new Font("微软雅黑", Font.PLAIN, 12);

        // 3. 顶部标题面板
        JPanel topPanel = new JPanel(new BorderLayout());
        topPanel.setOpaque(false);
        JLabel titleLabel = new JLabel("文档管理系统", SwingConstants.CENTER);
        titleLabel.setFont(titleFont);
        // 空白标签实现上下间距
        topPanel.add(new JLabel(" "), BorderLayout.NORTH);
        topPanel.add(titleLabel, BorderLayout.CENTER);
        topPanel.add(new JLabel(" "), BorderLayout.SOUTH);
        f.add(topPanel, BorderLayout.NORTH);

        // 4. 中间功能面板
        JPanel mainPanel = new JPanel(new BorderLayout());
        mainPanel.setOpaque(false);
        // 空白标签实现左右间距
        mainPanel.add(new JLabel("   "), BorderLayout.WEST);
        mainPanel.add(new JLabel("   "), BorderLayout.EAST);

        // 功能按钮：上传、生成目录、检查结构
        JPanel btnPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 15, 10));
        btnPanel.setOpaque(false);
        JButton btn1 = new JButton("上传文档");
        JButton btn2 = new JButton("生成目录");
        JButton btn3 = new JButton("检查结构");
        // 统一按钮大小，界面更整齐
        Dimension btnSize = new Dimension(120, 40);
        btn1.setPreferredSize(btnSize);
        btn2.setPreferredSize(btnSize);
        btn3.setPreferredSize(btnSize);
        btn1.setFont(btnFont);
        btn2.setFont(btnFont);
        btn3.setFont(btnFont);
        btnPanel.add(btn1);
        btnPanel.add(btn2);
        btn3.add(btn3);

        // 搜索功能：输入框 + 搜索按钮
        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 10, 10));
        searchPanel.setOpaque(false);
        JTextField tf = new JTextField(15);
        tf.setPreferredSize(new Dimension(180, 35));
        tf.setFont(btnFont);
        JButton btn4 = new JButton("搜索");
        btn4.setPreferredSize(new Dimension(100, 35));
        btn4.setFont(btnFont);
        searchPanel.add(tf);
        searchPanel.add(btn4);

        // 组合功能区域
        JPanel centerPanel = new JPanel();
        centerPanel.setLayout(new BoxLayout(centerPanel, BoxLayout.Y_AXIS));
        centerPanel.setOpaque(false);
        centerPanel.add(btnPanel);
        centerPanel.add(Box.createVerticalStrut(15));
        centerPanel.add(searchPanel);
        mainPanel.add(centerPanel, BorderLayout.CENTER);
        f.add(mainPanel, BorderLayout.CENTER);

        // 5. 底部提示栏：显示支持的文档格式
        JPanel bottomPanel = new JPanel(new BorderLayout());
        bottomPanel.setOpaque(false);
        JLabel supportTip = new JLabel("支持格式：doc / docx / pdf / txt", SwingConstants.CENTER);
        supportTip.setFont(tipFont);
        supportTip.setForeground(Color.GRAY);
        bottomPanel.add(supportTip, BorderLayout.CENTER);
        bottomPanel.add(new JLabel(" "), BorderLayout.SOUTH);
        f.add(bottomPanel, BorderLayout.SOUTH);

        // 6. 按钮绑定功能（点击按钮调用对应方法）
        btn1.addActionListener(e -> uploadFile());
        btn2.addActionListener(e -> buildToc());
        btn3.addActionListener(e -> checkStruct());
        btn4.addActionListener(e -> search(tf.getText()));

        // 显示主界面
        f.setVisible(true);
    }

    // ========================= 实体类：标题信息 =========================
    // 作用：封装标题的 层级 和 内容，方便数据传输和存储
    static class ParagraphInfo {
        int level;      // 标题层级（1/2）
        String content; // 标题文本内容

        // 构造方法：创建对象时赋值
        public ParagraphInfo(int level, String content) {
            this.level = level;
            this.content = content;
        }
    }
}