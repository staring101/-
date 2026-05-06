-- 1. 创建数据库（如果不存在）
CREATE DATABASE IF NOT EXISTS doc_manager;
USE doc_manager;

-- 2. 创建文档主表（先删后建，避免重复创建报错）
DROP TABLE IF EXISTS doc_section; -- 先删子表，否则外键会阻止父表删除
DROP TABLE IF EXISTS document;
CREATE TABLE document (
    id INT PRIMARY KEY AUTO_INCREMENT,
    filename VARCHAR(255) NOT NULL,  -- 文件名
    filepath VARCHAR(500) NOT NULL, -- 文件路径
    upload_time DATETIME DEFAULT CURRENT_TIMESTAMP -- 上传时间
);

-- 3. 创建章节信息表
CREATE TABLE doc_section (
    id INT PRIMARY KEY AUTO_INCREMENT,
    doc_id INT NOT NULL,  -- 关联文档ID
    level INT NOT NULL,   -- 标题层级（1/2级）
    content TEXT NOT NULL,-- 标题内容
    -- 补全外键定义，添加级联删除
    FOREIGN KEY (doc_id) REFERENCES document(id) ON DELETE CASCADE
);